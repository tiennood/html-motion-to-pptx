"""
bezier_math.py - 4-Bezier Closed Ellipse & Perspective Motion Path Generator
Used by html-motion-to-pptx skill to generate native PowerPoint SMIL motion paths.
"""

import math

KAPPA = (4.0 / 3.0) * (math.sqrt(2.0) - 1.0)  # 0.5522847498

def get_bezier_ellipse_path(cx: float, cy: float, a: float, b: float, 
                            init_x: float, init_y: float, theta0: float,
                            slide_w: float = 960.0, slide_h: float = 540.0) -> str:
    """
    Computes a mathematically closed 4-Bezier elliptical motion path in PowerPoint
    normalized coordinate space [0..1] relative to the starting position.
    
    :param cx: Center X of ellipse on slide
    :param cy: Center Y of ellipse on slide
    :param a: Semi-major radius (horizontal)
    :param b: Semi-minor radius (vertical, foreshortened by perspective)
    :param init_x: Initial shape X position
    :param init_y: Initial shape Y position
    :param theta0: Starting angle in radians
    :param slide_w: Slide width in points (default 960.0 for 16:9 widescreen)
    :param slide_h: Slide height in points (default 540.0 for 16:9 widescreen)
    :return: PowerPoint VML/SVG path string "M 0 0 C ... Z"
    """
    quarters = []
    for q in range(4):
        th_a = theta0 + (q * math.pi / 2.0)
        th_b = theta0 + ((q + 1) * math.pi / 2.0)

        # Control Point 1
        c1x = (cx + a * math.cos(th_a) - KAPPA * a * math.sin(th_a) - init_x) / slide_w
        c1y = (cy + b * math.sin(th_a) + KAPPA * b * math.cos(th_a) - init_y) / slide_h

        # Control Point 2
        c2x = (cx + a * math.cos(th_b) + KAPPA * a * math.sin(th_b) - init_x) / slide_w
        c2y = (cy + b * math.sin(th_b) - KAPPA * b * math.cos(th_b) - init_y) / slide_h

        # End of Quarter Arc
        end_x = (cx + a * math.cos(th_b) - init_x) / slide_w
        end_y = (cy + b * math.sin(th_b) - init_y) / slide_h

        quarters.append(f"C {c1x:.5f} {c1y:.5f} {c2x:.5f} {c2y:.5f} {end_x:.5f} {end_y:.5f}")

    return "M 0 0 " + " ".join(quarters) + " Z"


def project_3d_to_perspective(radius_3d: float, inclination_deg: float = 65.0) -> tuple[float, float]:
    """
    Maps 3D orbital radius to 2D elliptical axes (a, b) based on viewer inclination.
    """
    rad = math.radians(inclination_deg)
    a = radius_3d
    b = radius_3d * math.cos(rad)
    return a, b
