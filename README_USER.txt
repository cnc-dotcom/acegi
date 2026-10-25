RM CutPlan Studio v0.2.0 — User Notes

This is an early personal-use MVP for rectangular plate nesting and structural stock-length optimization.

IMPORTANT LIMITATIONS
- Plate parts are nested as rectangles; irregular DXF contours are not true-shape nested.
- DXF closed polylines/circles are converted to bounding rectangles; DXF units are assumed to be millimetres.
- The length optimizer is heuristic; verify kerf, trim, dimensions, and all results before production.
- It does not generate machine-ready CNC G-code/NC files.

The installed application bundles its Python runtime and does not require Python to be installed separately on the target PC.
