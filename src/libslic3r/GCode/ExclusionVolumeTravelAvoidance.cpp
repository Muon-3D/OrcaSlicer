#include "ExclusionVolumeTravelAvoidance.hpp"

#include "libslic3r/ClipperUtils.hpp"
#include "libslic3r/GCode/ExclusionVolumePathCheck.hpp"
#include "libslic3r/Line.hpp"
#include "libslic3r/Polygon.hpp"

#include <algorithm>
#include <cmath>

namespace Slic3r {

namespace {

static constexpr double Z_EPSILON_MM = 1e-6;
static constexpr double T_EPSILON = 1e-9;
static constexpr size_t MAX_REROUTE_ITERATIONS = 16;

struct SegmentIntersection
{
    double t { 0.0 };
    Point  point;
    size_t edge_idx { 0 };
};

static bool z_ranges_overlap(const double a_min, const double a_max, const double b_min, const double b_max)
{
    return a_max >= b_min - Z_EPSILON_MM && a_min <= b_max + Z_EPSILON_MM;
}

static bool line_intersection_t(
    const Point &a,
    const Point &b,
    const Point &c,
    const Point &d,
    double      &t,
    Point       &intersection)
{
    const Vec2d p = a.cast<double>();
    const Vec2d r = (b - a).cast<double>();
    const Vec2d q = c.cast<double>();
    const Vec2d s = (d - c).cast<double>();
    const double denom = cross2(r, s);
    if (std::abs(denom) < T_EPSILON)
        return false;

    const Vec2d qp = q - p;
    const double candidate_t = cross2(qp, s) / denom;
    const double u = cross2(qp, r) / denom;
    if (candidate_t < -T_EPSILON || candidate_t > 1.0 + T_EPSILON || u < -T_EPSILON || u > 1.0 + T_EPSILON)
        return false;

    t = std::clamp(candidate_t, 0.0, 1.0);
    intersection = (p + t * r).cast<coord_t>();
    return true;
}

template<typename Fn>
static void for_each_expolygon_contour(const ExPolygon &expolygon, Fn &&fn)
{
    fn(expolygon.contour);
    for (const Polygon &hole : expolygon.holes)
        fn(hole);
}

static void add_unique_param(std::vector<double> &params, const double t)
{
    const double clamped = std::clamp(t, 0.0, 1.0);
    if (std::none_of(params.begin(), params.end(), [clamped](double value) { return std::abs(value - clamped) < T_EPSILON; }))
        params.emplace_back(clamped);
}

static std::vector<SegmentIntersection> contour_intersections(const Point &a, const Point &b, const Polygon &polygon)
{
    std::vector<SegmentIntersection> intersections;
    if (polygon.points.size() < 2)
        return intersections;

    intersections.reserve(4);
    for (size_t edge_idx = 0; edge_idx < polygon.points.size(); ++edge_idx) {
        const Point &c = polygon.points[edge_idx];
        const Point &d = polygon.points[(edge_idx + 1) % polygon.points.size()];
        double t = 0.0;
        Point p;
        if (line_intersection_t(a, b, c, d, t, p)) {
            if (std::none_of(intersections.begin(), intersections.end(), [t](const SegmentIntersection &other) {
                    return std::abs(other.t - t) < T_EPSILON;
                }))
                intersections.push_back({ t, p, edge_idx });
        }
    }

    std::sort(intersections.begin(), intersections.end(), [](const SegmentIntersection &lhs, const SegmentIntersection &rhs) {
        return lhs.t < rhs.t;
    });
    return intersections;
}

static bool point_in_expolygons(const ExPolygons &expolygons, const Point &point, const bool border_result)
{
    return std::any_of(expolygons.begin(), expolygons.end(), [&point, border_result](const ExPolygon &expolygon) {
        return expolygon.contains(point, border_result);
    });
}

static std::vector<double> boundary_params_for_segment(const Point &a, const Point &b, const ExPolygons &expolygons)
{
    std::vector<double> params { 0.0, 1.0 };
    for (const ExPolygon &expolygon : expolygons) {
        for_each_expolygon_contour(expolygon, [&](const Polygon &contour) {
            for (const SegmentIntersection &intersection : contour_intersections(a, b, contour))
                add_unique_param(params, intersection.t);
        });
    }

    std::sort(params.begin(), params.end());
    return params;
}

static Point interpolate(const Point &a, const Point &b, const double t)
{
    return (a.cast<double>() + t * (b - a).cast<double>()).cast<coord_t>();
}

static bool segment_enters_expolygons_interior(const Point &a, const Point &b, const ExPolygons &expolygons)
{
    if (expolygons.empty())
        return false;
    if (point_in_expolygons(expolygons, a, false) || point_in_expolygons(expolygons, b, false))
        return true;

    const std::vector<double> params = boundary_params_for_segment(a, b, expolygons);
    for (size_t idx = 1; idx < params.size(); ++idx) {
        const double lo = params[idx - 1];
        const double hi = params[idx];
        if (hi - lo < T_EPSILON)
            continue;
        const Point mid = interpolate(a, b, 0.5 * (lo + hi));
        if (point_in_expolygons(expolygons, mid, false))
            return true;
    }

    return false;
}

static bool segment_inside_bed(const Point &a, const Point &b, const ExPolygons &valid_bed)
{
    if (valid_bed.empty())
        return true;
    if (!point_in_expolygons(valid_bed, a, true) || !point_in_expolygons(valid_bed, b, true))
        return false;

    const std::vector<double> params = boundary_params_for_segment(a, b, valid_bed);
    for (size_t idx = 1; idx < params.size(); ++idx) {
        const double lo = params[idx - 1];
        const double hi = params[idx];
        if (hi - lo < T_EPSILON)
            continue;
        const Point mid = interpolate(a, b, 0.5 * (lo + hi));
        if (!point_in_expolygons(valid_bed, mid, false))
            return false;
    }

    return true;
}

static double polyline_length(const Polyline &polyline)
{
    double length = 0.0;
    for (size_t idx = 1; idx < polyline.points.size(); ++idx)
        length += (polyline.points[idx] - polyline.points[idx - 1]).cast<double>().norm();
    return length;
}

static void append_point(Points &points, const Point &point)
{
    if (points.empty() || points.back() != point)
        points.emplace_back(point);
}

static Polyline make_forward_detour(
    const Point               &a,
    const Point               &b,
    const Polygon             &contour,
    const SegmentIntersection &entry,
    const SegmentIntersection &exit)
{
    Points points;
    append_point(points, a);
    append_point(points, entry.point);

    const size_t n = contour.points.size();
    for (size_t idx = (entry.edge_idx + 1) % n; idx != (exit.edge_idx + 1) % n; idx = (idx + 1) % n)
        append_point(points, contour.points[idx]);

    append_point(points, exit.point);
    append_point(points, b);
    return Polyline(std::move(points));
}

static Polyline make_backward_detour(
    const Point               &a,
    const Point               &b,
    const Polygon             &contour,
    const SegmentIntersection &entry,
    const SegmentIntersection &exit)
{
    Points points;
    append_point(points, a);
    append_point(points, entry.point);

    const size_t n = contour.points.size();
    const size_t last_vertex = (exit.edge_idx + 1) % n;
    for (size_t idx = entry.edge_idx;; idx = (idx == 0) ? n - 1 : idx - 1) {
        append_point(points, contour.points[idx]);
        if (idx == last_vertex)
            break;
    }

    append_point(points, exit.point);
    append_point(points, b);
    return Polyline(std::move(points));
}

static bool polyline_inside_bed(const Polyline &polyline, const ExPolygons &valid_bed)
{
    for (size_t idx = 1; idx < polyline.points.size(); ++idx)
        if (!segment_inside_bed(polyline.points[idx - 1], polyline.points[idx], valid_bed))
            return false;
    return true;
}

static bool polyline_intersects_obstacles(const Polyline &polyline, const ExPolygons &obstacles)
{
    for (size_t idx = 1; idx < polyline.points.size(); ++idx)
        if (segment_enters_expolygons_interior(polyline.points[idx - 1], polyline.points[idx], obstacles))
            return true;
    return false;
}

static std::optional<size_t> first_intersected_obstacle(const Point &a, const Point &b, const ExPolygons &obstacles)
{
    for (size_t idx = 0; idx < obstacles.size(); ++idx) {
        ExPolygons single { obstacles[idx] };
        if (segment_enters_expolygons_interior(a, b, single))
            return idx;
    }

    return std::nullopt;
}

static std::optional<Polyline> detour_around_obstacle(
    const Point      &a,
    const Point      &b,
    const ExPolygon  &obstacle,
    const ExPolygons &valid_bed)
{
    const std::vector<SegmentIntersection> intersections = contour_intersections(a, b, obstacle.contour);
    if (intersections.size() < 2)
        return std::nullopt;

    const SegmentIntersection &entry = intersections.front();
    const SegmentIntersection &exit  = intersections.back();
    Polyline forward  = make_forward_detour(a, b, obstacle.contour, entry, exit);
    Polyline backward = make_backward_detour(a, b, obstacle.contour, entry, exit);

    const bool forward_valid = polyline_inside_bed(forward, valid_bed);
    const bool backward_valid = polyline_inside_bed(backward, valid_bed);
    if (!forward_valid && !backward_valid)
        return std::nullopt;
    if (forward_valid && !backward_valid)
        return forward;
    if (!forward_valid && backward_valid)
        return backward;

    return polyline_length(forward) <= polyline_length(backward) ? forward : backward;
}

static void replace_segment(Polyline &path, const size_t segment_idx, const Polyline &replacement)
{
    Points points;
    points.reserve(path.points.size() + replacement.points.size());
    for (size_t idx = 0; idx < segment_idx; ++idx)
        append_point(points, path.points[idx]);
    for (const Point &point : replacement.points)
        append_point(points, point);
    for (size_t idx = segment_idx + 2; idx < path.points.size(); ++idx)
        append_point(points, path.points[idx]);
    path.points = std::move(points);
}

static Polyline simplify_path(const Polyline &path, const ExPolygons &obstacles, const ExPolygons &valid_bed)
{
    if (path.points.size() <= 2)
        return path;

    Points simplified;
    simplified.reserve(path.points.size());
    size_t current = 0;
    append_point(simplified, path.points.front());
    while (current + 1 < path.points.size()) {
        size_t next = path.points.size() - 1;
        for (; next > current + 1; --next) {
            if (segment_inside_bed(path.points[current], path.points[next], valid_bed) &&
                !segment_enters_expolygons_interior(path.points[current], path.points[next], obstacles))
                break;
        }
        append_point(simplified, path.points[next]);
        current = next;
    }

    return Polyline(std::move(simplified));
}

} // namespace

void ExclusionVolumeTravelAvoidance::init(const PrintConfig &config, const Vec3d &plate_origin)
{
    this->clear();
    m_regions = translated_bed_exclusion_regions(config, plate_origin);

    m_bed_shape.points = get_bed_shape(config);
    const Point origin(scale_(plate_origin.x()), scale_(plate_origin.y()));
    m_bed_shape.translate(origin);
    m_bed_shape.make_counter_clockwise();

    double max_nozzle_diameter = 0.0;
    for (double diameter : config.nozzle_diameter.values)
        max_nozzle_diameter = std::max(max_nozzle_diameter, diameter);

    // Keep the detour visibly outside the configured region without inventing a
    // large, hidden clearance policy. The post-generation path check still uses
    // the exact configured volume as the source of truth.
    m_clearance = std::max<coord_t>(SCALED_EPSILON, coord_t(scale_(std::max(0.05, 0.1 * max_nozzle_diameter))));
}

void ExclusionVolumeTravelAvoidance::clear()
{
    m_regions.clear();
    m_bed_shape.points.clear();
    m_clearance = SCALED_EPSILON;
}

std::optional<ExclusionVolumeTravelAvoidance::ActiveObstacles> ExclusionVolumeTravelAvoidance::active_obstacles(
    const double z_min,
    const double z_max) const
{
    Polygons active_polygons;
    active_polygons.reserve(m_regions.size());
    for (const BedExcludeRegion &region : m_regions)
        if (z_ranges_overlap(z_min, z_max, region.z_min, region.z_max))
            active_polygons.emplace_back(region.polygon);

    if (active_polygons.empty())
        return std::nullopt;

    ActiveObstacles active;
    active.obstacles = union_ex(expand(active_polygons, float(m_clearance)));
    if (m_bed_shape.points.size() >= 3) {
        active.valid_bed = offset_ex(m_bed_shape, -float(m_clearance));
        if (active.valid_bed.empty())
            active.valid_bed = union_ex(Polygons { m_bed_shape });
    }

    return active;
}

ExclusionVolumeTravelAvoidance::Result ExclusionVolumeTravelAvoidance::route(
    const Polyline &travel,
    const double    start_z,
    const double    end_z) const
{
    Result result;
    result.path = travel;
    if (travel.points.size() < 2 || m_regions.empty())
        return result;

    const double z_min = std::min(start_z, end_z);
    const double z_max = std::max(start_z, end_z);
    std::optional<ActiveObstacles> active = this->active_obstacles(z_min, z_max);
    if (!active || active->obstacles.empty())
        return result;

    const Point &start = travel.points.front();
    const Point &end   = travel.points.back();
    if (point_in_expolygons(active->obstacles, start, false) || point_in_expolygons(active->obstacles, end, false)) {
        result.status = Status::EndpointInside;
        return result;
    }

    bool needs_reroute = false;
    for (size_t idx = 1; idx < travel.points.size(); ++idx) {
        if (segment_enters_expolygons_interior(travel.points[idx - 1], travel.points[idx], active->obstacles)) {
            needs_reroute = true;
            break;
        }
    }
    if (!needs_reroute)
        return result;

    Polyline path = travel;
    bool rerouted = false;
    for (size_t iteration = 0; iteration < MAX_REROUTE_ITERATIONS; ++iteration) {
        bool changed = false;
        for (size_t segment_idx = 0; segment_idx + 1 < path.points.size(); ++segment_idx) {
            const Point &a = path.points[segment_idx];
            const Point &b = path.points[segment_idx + 1];
            if (!segment_inside_bed(a, b, active->valid_bed)) {
                result.status = Status::Failed;
                return result;
            }

            std::optional<size_t> obstacle_idx = first_intersected_obstacle(a, b, active->obstacles);
            if (!obstacle_idx)
                continue;

            std::optional<Polyline> detour = detour_around_obstacle(a, b, active->obstacles[*obstacle_idx], active->valid_bed);
            if (!detour) {
                result.status = Status::Failed;
                return result;
            }

            replace_segment(path, segment_idx, *detour);
            changed = true;
            rerouted = true;
            break;
        }

        if (!changed) {
            path = simplify_path(path, active->obstacles, active->valid_bed);
            if (!polyline_inside_bed(path, active->valid_bed) || polyline_intersects_obstacles(path, active->obstacles)) {
                result.status = Status::Failed;
                return result;
            }
            result.path = std::move(path);
            result.status = rerouted ? Status::Rerouted : Status::Unchanged;
            return result;
        }
    }

    result.status = Status::Failed;
    return result;
}

} // namespace Slic3r
