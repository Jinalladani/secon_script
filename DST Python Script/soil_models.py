"""
DST Synthetic Data Generator
Stage 1 - Soil Behaviour Models

This module contains the finalized Version 1.0 behaviour
profiles for the 8 soil types.
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class SoilProfile:
    name: str

    # Typical engineering ranges
    c_min: float
    c_max: float
    phi_min: float
    phi_max: float

    # Curve behaviour
    peak_min: float
    peak_max: float

    # 0 = very broad, 1 = very sharp
    peak_sharpness: float

    # Expected post-peak reduction
    post_peak_min: float
    post_peak_max: float

    # Residual strength as percentage of peak
    residual_min: float
    residual_max: float

    # Observation noise
    noise_level: float

    # Local irregularity
    irregularity: float

    description: str


SOIL_PROFILES = {

    "Soft Clay": SoilProfile(
        name="Soft Clay",

        c_min=0.20,
        c_max=0.80,
        phi_min=5.0,
        phi_max=12.0,

        peak_min=0.60,
        peak_max=0.80,

        peak_sharpness=0.15,

        post_peak_min=0.00,
        post_peak_max=0.03,

        residual_min=0.97,
        residual_max=1.00,

        noise_level=0.008,
        irregularity=0.10,

        description="Smooth rise, broad plateau and no sharp peak."
    ),

    "Medium Clay": SoilProfile(
        name="Medium Clay",

        c_min=0.30,
        c_max=1.20,
        phi_min=10.0,
        phi_max=18.0,

        peak_min=0.50,
        peak_max=0.70,

        peak_sharpness=0.35,

        post_peak_min=0.03,
        post_peak_max=0.08,

        residual_min=0.92,
        residual_max=0.97,

        noise_level=0.012,
        irregularity=0.15,

        description="Moderate peak with limited post-peak softening."
    ),

    "Stiff Clay": SoilProfile(
        name="Stiff Clay",

        c_min=0.50,
        c_max=2.50,
        phi_min=15.0,
        phi_max=25.0,

        peak_min=0.40,
        peak_max=0.65,

        peak_sharpness=0.55,

        post_peak_min=0.02,
        post_peak_max=0.06,

        residual_min=0.94,
        residual_max=0.98,

        noise_level=0.010,
        irregularity=0.12,

        description="Strong peak with relatively little softening."
    ),

    "Silty Soil": SoilProfile(
        name="Silty Soil",

        c_min=0.05,
        c_max=0.40,
        phi_min=22.0,
        phi_max=30.0,

        peak_min=0.45,
        peak_max=0.70,

        peak_sharpness=0.45,

        post_peak_min=0.04,
        post_peak_max=0.10,

        residual_min=0.90,
        residual_max=0.96,

        noise_level=0.018,
        irregularity=0.30,

        description="Moderate peak with noticeable irregularity."
    ),

    "Loose Sand": SoilProfile(
        name="Loose Sand",

        c_min=0.00,
        c_max=0.10,
        phi_min=28.0,
        phi_max=34.0,

        peak_min=0.55,
        peak_max=0.80,

        peak_sharpness=0.20,

        post_peak_min=0.00,
        post_peak_max=0.05,

        residual_min=0.95,
        residual_max=1.00,

        noise_level=0.018,
        irregularity=0.25,

        description="Broad peak with little post-peak reduction."
    ),

    "Medium Dense Sand": SoilProfile(
        name="Medium Dense Sand",

        c_min=0.00,
        c_max=0.05,
        phi_min=34.0,
        phi_max=38.0,

        peak_min=0.35,
        peak_max=0.60,

        peak_sharpness=0.75,

        post_peak_min=0.05,
        post_peak_max=0.15,

        residual_min=0.85,
        residual_max=0.95,

        noise_level=0.018,
        irregularity=0.25,

        description="Clear and relatively sharp peak."
    ),

    "Dense Sand": SoilProfile(
        name="Dense Sand",

        c_min=0.00,
        c_max=0.02,
        phi_min=38.0,
        phi_max=45.0,

        peak_min=0.25,
        peak_max=0.50,

        peak_sharpness=0.95,

        post_peak_min=0.10,
        post_peak_max=0.25,

        residual_min=0.75,
        residual_max=0.90,

        noise_level=0.020,
        irregularity=0.30,

        description="Pronounced peak followed by significant softening."
    ),

    "Gravelly Sand": SoilProfile(
        name="Gravelly Sand",

        c_min=0.00,
        c_max=0.00,
        phi_min=40.0,
        phi_max=48.0,

        peak_min=0.30,
        peak_max=0.55,

        peak_sharpness=0.80,

        post_peak_min=0.08,
        post_peak_max=0.20,

        residual_min=0.80,
        residual_max=0.92,

        noise_level=0.030,
        irregularity=0.65,

        description="Irregular peak due to particle rearrangement/crushing."
    ),
}


def get_soil_profile(soil_name: str) -> SoilProfile:
    """Return the selected soil behaviour profile."""

    if soil_name not in SOIL_PROFILES:
        raise ValueError(f"Unknown soil type: {soil_name}")

    return SOIL_PROFILES[soil_name]


def get_soil_names():
    """Return soil names for the GUI."""

    return list(SOIL_PROFILES.keys())


def validate_target_strength(soil_name: str, cohesion: float, phi: float):
    """
    Check whether target c and phi fall within the typical
    range defined for the selected soil.

    Returns:
        (is_inside_range, message)
    """

    profile = get_soil_profile(soil_name)

    c_ok = profile.c_min <= cohesion <= profile.c_max
    phi_ok = profile.phi_min <= phi <= profile.phi_max

    if c_ok and phi_ok:
        return True, "Target c and φ are within the typical range."

    messages = []

    if not c_ok:
        messages.append(
            f"c = {cohesion:.3f} kg/cm² "
            f"(typical {profile.c_min:.2f}–{profile.c_max:.2f})"
        )

    if not phi_ok:
        messages.append(
            f"φ = {phi:.2f}° "
            f"(typical {profile.phi_min:.1f}–{profile.phi_max:.1f})"
        )

    return False, "Target strength outside typical range:\n" + "\n".join(messages)