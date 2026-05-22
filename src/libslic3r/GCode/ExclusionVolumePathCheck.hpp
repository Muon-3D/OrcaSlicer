#ifndef slic3r_GCode_ExclusionVolumePathCheck_hpp_
#define slic3r_GCode_ExclusionVolumePathCheck_hpp_

#include "libslic3r/libslic3r.h"
#include "libslic3r/PrintConfig.hpp"

#include <cstddef>
#include <optional>
#include <vector>

namespace Slic3r {

enum class EMoveType : unsigned char;
struct GCodeProcessorResult;

enum class ExclusionVolumeMoveClass : unsigned char
{
    Travel,
    Extrude,
    Other
};

struct ExclusionVolumePathHit
{
    size_t                   move_id { 0 };
    unsigned int             gcode_id { 0 };
    EMoveType                move_type {};
    ExclusionVolumeMoveClass move_class { ExclusionVolumeMoveClass::Other };
    size_t                   region_id { 0 };
    Vec3f                    from { Vec3f::Zero() };
    Vec3f                    to { Vec3f::Zero() };
};

struct ExclusionVolumePathCheckResult
{
    bool has_any_conflict { false };
    bool has_travel_conflict { false };
    bool has_extrusion_conflict { false };
    bool has_other_motion_conflict { false };
    std::optional<ExclusionVolumePathHit> first_hit;
};

std::vector<BedExcludeRegion> translated_bed_exclusion_regions(
    const PrintConfig &config,
    const Vec3d       &plate_origin);

ExclusionVolumePathCheckResult check_gcode_moves_against_exclusion_volumes(
    const GCodeProcessorResult        &result,
    const std::vector<BedExcludeRegion> &regions);

void apply_exclusion_volume_path_check_result(
    GCodeProcessorResult                   &result,
    const ExclusionVolumePathCheckResult   &check_result);

} // namespace Slic3r

#endif
