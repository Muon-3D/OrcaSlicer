#ifndef slic3r_GCode_ExclusionVolumeTravelAvoidance_hpp_
#define slic3r_GCode_ExclusionVolumeTravelAvoidance_hpp_

#include "libslic3r/ExPolygon.hpp"
#include "libslic3r/Polyline.hpp"
#include "libslic3r/PrintConfig.hpp"

#include <optional>
#include <vector>

namespace Slic3r {

class ExclusionVolumeTravelAvoidance
{
public:
    enum class Status : unsigned char
    {
        Unchanged,
        Rerouted,
        EndpointInside,
        Failed
    };

    struct Result
    {
        Status   status { Status::Unchanged };
        Polyline path;

        bool rerouted() const { return status == Status::Rerouted; }
    };

    void init(const PrintConfig &config, const Vec3d &plate_origin);
    void clear();

    bool empty() const { return m_regions.empty(); }

    // Input and output are scaled XY points in generated G-code coordinates
    // before the writer subtracts the plate offset.
    Result route(const Polyline &travel, double start_z, double end_z) const;

private:
    struct ActiveObstacles
    {
        ExPolygons obstacles;
        ExPolygons valid_bed;
    };

    std::optional<ActiveObstacles> active_obstacles(double z_min, double z_max) const;

    std::vector<BedExcludeRegion> m_regions;
    Polygon                       m_bed_shape;
    coord_t                       m_clearance { SCALED_EPSILON };
};

} // namespace Slic3r

#endif
