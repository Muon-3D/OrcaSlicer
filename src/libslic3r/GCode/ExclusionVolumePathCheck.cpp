#include "ExclusionVolumePathCheck.hpp"

#include "libslic3r/AABBTreeLines.hpp"
#include "libslic3r/BoundingBox.hpp"
#include "libslic3r/GCode/GCodeProcessor.hpp"
#include "libslic3r/Line.hpp"
#include "libslic3r/Polygon.hpp"

#include <algorithm>
#include <cmath>
#include <utility>

namespace Slic3r {

namespace {

static constexpr double Z_EPSILON_MM = 1e-6;

struct PreparedExclusionRegion
{
    BedExcludeRegion region;
    BoundingBox      bbox;
    Lines            edges;
    AABBTreeLines::LinesDistancer<Line> edge_tree;
};

static Point scaled_point(const Vec3f &point)
{
    return { scale_(double(point.x())), scale_(double(point.y())) };
}

static BoundingBox line_bbox(const Point &a, const Point &b)
{
    BoundingBox bbox;
    bbox.merge(a);
    bbox.merge(b);
    bbox.offset(SCALED_EPSILON);
    return bbox;
}

static bool intervals_overlap(coord_t a_min, coord_t a_max, coord_t b_min, coord_t b_max)
{
    if (a_min > a_max)
        std::swap(a_min, a_max);
    if (b_min > b_max)
        std::swap(b_min, b_max);
    return std::max(a_min, b_min) <= std::min(a_max, b_max);
}

static bool collinear_segments_overlap(const Line &a, const Line &b)
{
    const Vec2crd a_dir = a.b - a.a;
    if (a_dir == Vec2crd::Zero())
        return a.a == b.a || a.a == b.b;

    if (cross2(a_dir, b.a - a.a) != 0 || cross2(a_dir, b.b - a.a) != 0)
        return false;

    return std::abs(a_dir.x()) >= std::abs(a_dir.y()) ?
        intervals_overlap(a.a.x(), a.b.x(), b.a.x(), b.b.x()) :
        intervals_overlap(a.a.y(), a.b.y(), b.a.y(), b.b.y());
}

static bool overlaps_collinear_boundary(const Line &line, const PreparedExclusionRegion &region)
{
    const BoundingBox line_bounds = line_bbox(line.a, line.b);
    for (const Line &edge : region.edges) {
        if (!line_bounds.overlap(line_bbox(edge.a, edge.b)))
            continue;
        if (collinear_segments_overlap(line, edge))
            return true;
    }

    return false;
}

static ExclusionVolumeMoveClass move_class(const EMoveType type)
{
    if (type == EMoveType::Travel)
        return ExclusionVolumeMoveClass::Travel;
    if (type == EMoveType::Extrude || type == EMoveType::Wipe)
        return ExclusionVolumeMoveClass::Extrude;
    return ExclusionVolumeMoveClass::Other;
}

static bool is_physical_motion_type(const EMoveType type)
{
    // Marker records such as seam, color change, pause and custom G-code do not
    // represent physical line segments in the rendered toolpath.
    switch (type) {
    case EMoveType::Travel:
    case EMoveType::Wipe:
    case EMoveType::Extrude:
    case EMoveType::Retract:
    case EMoveType::Unretract:
        return true;
    default:
        return false;
    }
}

static void record_hit(
    ExclusionVolumePathCheckResult &result,
    const size_t                    move_id,
    const GCodeProcessorResult::MoveVertex &move,
    const size_t                    region_id,
    const Vec3f                    &from,
    const Vec3f                    &to)
{
    const ExclusionVolumeMoveClass cls = move_class(move.type);
    result.has_any_conflict = true;
    result.has_travel_conflict |= cls == ExclusionVolumeMoveClass::Travel;
    result.has_extrusion_conflict |= cls == ExclusionVolumeMoveClass::Extrude;
    result.has_other_motion_conflict |= cls == ExclusionVolumeMoveClass::Other;

    if (!result.first_hit) {
        result.first_hit = ExclusionVolumePathHit{
            move_id,
            move.gcode_id,
            move.type,
            cls,
            region_id,
            from,
            to
        };
    }
}

static bool clip_segment_to_z_range(
    const Vec3f  &from,
    const Vec3f  &to,
    const double  z_min,
    const double  z_max,
    Vec3f        &out_from,
    Vec3f        &out_to)
{
    const double from_z = double(from.z());
    const double to_z   = double(to.z());
    const double seg_min_z = std::min(from_z, to_z);
    const double seg_max_z = std::max(from_z, to_z);

    if (seg_max_z < z_min - Z_EPSILON_MM || seg_min_z > z_max + Z_EPSILON_MM)
        return false;

    double t0 = 0.0;
    double t1 = 1.0;
    const double dz = to_z - from_z;

    if (std::abs(dz) < Z_EPSILON_MM) {
        if (from_z < z_min - Z_EPSILON_MM || from_z > z_max + Z_EPSILON_MM)
            return false;
    } else {
        const double t_min = (z_min - from_z) / dz;
        const double t_max = (z_max - from_z) / dz;
        t0 = std::max(t0, std::min(t_min, t_max));
        t1 = std::min(t1, std::max(t_min, t_max));
        if (t0 > t1 + Z_EPSILON_MM)
            return false;
    }

    t0 = std::clamp(t0, 0.0, 1.0);
    t1 = std::clamp(t1, 0.0, 1.0);
    const Vec3f delta = to - from;
    out_from = from + float(t0) * delta;
    out_to   = from + float(t1) * delta;
    return true;
}

static std::vector<PreparedExclusionRegion> prepare_regions(const std::vector<BedExcludeRegion> &regions)
{
    std::vector<PreparedExclusionRegion> prepared;
    prepared.reserve(regions.size());

    for (const BedExcludeRegion &region : regions) {
        if (region.polygon.points.size() < 3)
            continue;
        if (region.z_max < region.z_min)
            continue;

        PreparedExclusionRegion item;
        item.region = region;
        item.bbox = BoundingBox(region.polygon.points);
        item.bbox.offset(SCALED_EPSILON);
        item.edges = to_lines(region.polygon);
        item.edge_tree = AABBTreeLines::LinesDistancer<Line>(item.edges);
        prepared.emplace_back(std::move(item));
    }

    return prepared;
}

static bool segment_intersects_region(
    const Vec3f                    &from,
    const Vec3f                    &to,
    const PreparedExclusionRegion  &region)
{
    Vec3f clipped_from;
    Vec3f clipped_to;
    if (!clip_segment_to_z_range(from, to, region.region.z_min, region.region.z_max, clipped_from, clipped_to))
        return false;

    const Point a = scaled_point(clipped_from);
    const Point b = scaled_point(clipped_to);
    const BoundingBox segment_bbox = line_bbox(a, b);
    if (!segment_bbox.overlap(region.bbox))
        return false;

    if (contains(region.region.polygon, a, true) || contains(region.region.polygon, b, true))
        return true;

    if (a == b)
        return false;

    const Line line(a, b);
    if (!region.edge_tree.intersections_with_line<false>(line).empty())
        return true;

    // The AABB line-tree intersection intentionally ignores collinear overlap.
    return overlaps_collinear_boundary(line, region);
}

static bool segment_intersects_any_region(
    const Vec3f                                &from,
    const Vec3f                                &to,
    const std::vector<PreparedExclusionRegion> &regions,
    size_t                                     &region_id)
{
    for (size_t idx = 0; idx < regions.size(); ++idx) {
        if (segment_intersects_region(from, to, regions[idx])) {
            region_id = idx;
            return true;
        }
    }

    return false;
}

template<typename Fn>
static void for_each_move_segment(
    const GCodeProcessorResult::MoveVertex &prev,
    const GCodeProcessorResult::MoveVertex &curr,
    Fn                                      fn)
{
    const size_t interpolation_points = curr.is_arc_move_with_interpolation_points() ?
        curr.interpolation_points.size() : 0;

    Vec3f from = prev.position;
    for (size_t idx = 0; idx <= interpolation_points; ++idx) {
        const Vec3f &to = idx == interpolation_points ? curr.position : curr.interpolation_points[idx];
        fn(from, to);
        from = to;
    }
}

} // namespace

std::vector<BedExcludeRegion> translated_bed_exclusion_regions(
    const PrintConfig &config,
    const Vec3d       &plate_origin)
{
    std::vector<BedExcludeRegion> regions = get_bed_excluded_regions(config);
    const Point origin(scale_(plate_origin.x()), scale_(plate_origin.y()));

    for (BedExcludeRegion &region : regions)
        region.polygon.translate(origin);

    return regions;
}

ExclusionVolumePathCheckResult check_gcode_moves_against_exclusion_volumes(
    const GCodeProcessorResult          &result,
    const std::vector<BedExcludeRegion> &regions)
{
    ExclusionVolumePathCheckResult check_result;
    const std::vector<PreparedExclusionRegion> prepared_regions = prepare_regions(regions);
    if (prepared_regions.empty() || result.moves.size() < 2)
        return check_result;

    for (size_t move_id = 1; move_id < result.moves.size(); ++move_id) {
        const GCodeProcessorResult::MoveVertex &prev = result.moves[move_id - 1];
        const GCodeProcessorResult::MoveVertex &curr = result.moves[move_id];
        if (!is_physical_motion_type(curr.type))
            continue;

        for_each_move_segment(prev, curr, [&](const Vec3f &from, const Vec3f &to) {
            if (check_result.has_any_conflict)
                return;

            size_t region_id = 0;
            if (segment_intersects_any_region(from, to, prepared_regions, region_id))
                record_hit(check_result, move_id, curr, region_id, from, to);
        });

        if (check_result.has_any_conflict)
            break;
    }

    return check_result;
}

void apply_exclusion_volume_path_check_result(
    GCodeProcessorResult                 &result,
    const ExclusionVolumePathCheckResult &check_result)
{
    result.exclusion_volume_path_checked = true;
    result.exclusion_volume_path_conflict = check_result.has_any_conflict;
    result.exclusion_volume_travel_conflict = check_result.has_travel_conflict;
    result.exclusion_volume_extrusion_conflict = check_result.has_extrusion_conflict;
    result.exclusion_volume_other_motion_conflict = check_result.has_other_motion_conflict;
    result.exclusion_volume_conflict_gcode_id = check_result.first_hit ? check_result.first_hit->gcode_id : 0;
    result.exclusion_volume_conflict_move_id = check_result.first_hit ? check_result.first_hit->move_id : 0;
    result.exclusion_volume_conflict_move_type = check_result.first_hit ? check_result.first_hit->move_type : EMoveType::Noop;
}

} // namespace Slic3r
