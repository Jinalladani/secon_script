"""
DST Synthetic Data Generator
Stage 2 - Final Behaviour Model

No Excel writing in this stage.

Model components:
    - 61 displacement observations
    - soil-specific initial response
    - soil/density-dependent peak displacement
    - independent trial-to-trial peak-strength variation
    - post-peak softening
    - residual strength
    - controlled natural + instrument/logger imperfections
    - Mohr-Coulomb strength-envelope calculation
    - controlled regeneration/QC for c, phi and R²

This is an engineering-based synthetic generator.
It is NOT calibrated against an actual laboratory database.
"""

import math
import random
from dataclasses import dataclass

from soil_models import get_soil_profile


# ============================================================
# DATA STRUCTURE
# ============================================================

@dataclass
class DSTTrial:
    normal_stress: float
    displacement: list
    shear_stress: list

    failure_shear_stress: float
    peak_shear_stress: float
    peak_displacement: float

    residual_shear_stress: float


# ============================================================
# BASIC FUNCTIONS
# ============================================================

def mohr_coulomb_failure(c, phi, normal_stress):
    """
    Theoretical Mohr-Coulomb shear strength.

    c     : kg/cm²
    phi   : degrees
    sigma : kg/cm²
    """
    return (
        c
        + normal_stress
        * math.tan(math.radians(phi))
    )


def clamp(value, minimum, maximum):
    return max(minimum, min(value, maximum))


# ============================================================
# ACTUAL STRENGTH ENVELOPE
# ============================================================

def calculate_strength_envelope(trials):
    """
    Calculate c, phi and R² from the generated peak stresses.

    tau = c + sigma * tan(phi)
    """

    if len(trials) < 2:
        raise ValueError(
            "At least two trials are required."
        )

    sigma = [
        t.normal_stress
        for t in trials
    ]

    tau = [
        t.peak_shear_stress
        for t in trials
    ]

    sigma_mean = sum(sigma) / len(sigma)
    tau_mean = sum(tau) / len(tau)

    denominator = sum(
        (x - sigma_mean) ** 2
        for x in sigma
    )

    if denominator <= 0:
        raise ValueError(
            "Normal stresses must not be identical."
        )

    numerator = sum(
        (x - sigma_mean) * (y - tau_mean)
        for x, y in zip(sigma, tau)
    )

    slope = numerator / denominator

    cohesion = (
        tau_mean
        - slope * sigma_mean
    )

    friction_angle = math.degrees(
        math.atan(slope)
    )

    ss_total = sum(
        (y - tau_mean) ** 2
        for y in tau
    )

    ss_residual = sum(
        (
            y
            - (
                cohesion
                + slope * x
            )
        ) ** 2
        for x, y in zip(sigma, tau)
    )

    if ss_total > 0:
        r_squared = (
            1.0
            - ss_residual / ss_total
        )
    else:
        r_squared = 1.0

    return {
        "cohesion": cohesion,
        "phi": friction_angle,
        "r_squared": r_squared,
    }


# ============================================================
# DENSITY / CONSISTENCY CLASS
# ============================================================

def classify_density(soil_type, density):
    """
    Approximate behaviour classification used only by the
    synthetic curve model.
    """

    soil = soil_type.lower()

    if "sand" in soil:

        if density < 1.60:
            return "Loose"

        elif density < 1.72:
            return "Medium Dense"

        elif density < 1.82:
            return "Dense"

        else:
            return "Very Dense"

    if "clay" in soil:

        if density < 1.45:
            return "Soft"

        elif density < 1.55:
            return "Medium"

        elif density < 1.70:
            return "Stiff"

        else:
            return "Very Stiff"

    if "silty" in soil:

        if density < 1.50:
            return "Loose"

        elif density < 1.65:
            return "Medium"

        else:
            return "Dense"

    return "Medium"


# ============================================================
# DENSITY BEHAVIOUR
# ============================================================

def density_behaviour_factor(
    soil_type,
    density_class
):
    """
    Modifiers for peak location, peak sharpness and softening.
    """

    soil = soil_type.lower()

    if "sand" in soil:

        if density_class == "Loose":
            return {
                "peak_shift": 1.15,
                "sharpness": 0.75,
                "softening": 0.65,
            }

        if density_class == "Medium Dense":
            return {
                "peak_shift": 1.00,
                "sharpness": 1.00,
                "softening": 1.00,
            }

        if density_class == "Dense":
            return {
                "peak_shift": 0.85,
                "sharpness": 1.20,
                "softening": 1.15,
            }

        return {
            "peak_shift": 0.75,
            "sharpness": 1.30,
            "softening": 1.25,
        }

    if "clay" in soil:

        if density_class == "Soft":
            return {
                "peak_shift": 1.15,
                "sharpness": 0.75,
                "softening": 0.70,
            }

        if density_class == "Medium":
            return {
                "peak_shift": 1.00,
                "sharpness": 1.00,
                "softening": 1.00,
            }

        if density_class == "Stiff":
            return {
                "peak_shift": 0.90,
                "sharpness": 1.10,
                "softening": 1.05,
            }

        return {
            "peak_shift": 0.82,
            "sharpness": 1.20,
            "softening": 1.10,
        }

    # Silty soil / fallback
    return {
        "peak_shift": 1.00,
        "sharpness": 1.00,
        "softening": 1.00,
    }


# ============================================================
# TRIAL STRENGTH VARIABILITY
# ============================================================

def strength_variation_range(
    soil_type,
    density_class
):
    """
    Independent peak-strength variability.

    The agreed nominal range is approximately ±2–4%.
    Soil/density controls where inside that range the
    variation is centred.

    This controls specimen variability, NOT logger noise.
    """

    soil = soil_type.lower()

    if "gravelly sand" in soil:
        base = 0.040

    elif "very dense sand" in soil:
        base = 0.035

    elif "dense sand" in soil:
        base = 0.035

    elif "medium dense sand" in soil:
        base = 0.030

    elif "loose sand" in soil:
        base = 0.040

    elif "soft clay" in soil:
        base = 0.035

    elif "medium clay" in soil:
        base = 0.040

    elif "stiff clay" in soil:
        base = 0.030

    elif "silty soil" in soil:
        base = 0.040

    else:
        base = 0.035

    d = density_class.lower()

    if "loose" in d:
        base += 0.002

    elif "very dense" in d:
        base -= 0.002

    return clamp(
        base,
        0.020,
        0.040
    )


# ============================================================
# ONE DST CURVE
# ============================================================

def generate_curve(
    soil_type,
    density,
    c,
    phi,
    normal_stress,
    seed,
    n_points=61,
    max_displacement=6.0,
    trial_strength_factor=1.0,
    strength_factor=None,
):
    """
    Generate one complete DST curve.

    trial_strength_factor is the preferred argument.

    strength_factor is retained as an alias so previous versions
    of the generator remain compatible.
    """

    if strength_factor is not None:
        trial_strength_factor = strength_factor

    rng = random.Random(seed)

    profile = get_soil_profile(
        soil_type
    )

    density_class = classify_density(
        soil_type,
        density
    )

    density_factor = density_behaviour_factor(
        soil_type,
        density_class
    )

    # --------------------------------------------------------
    # THEORETICAL FAILURE STRESS
    # --------------------------------------------------------

    theoretical_failure_tau = (
        mohr_coulomb_failure(
            c,
            phi,
            normal_stress
        )
    )

    # --------------------------------------------------------
    # TRIAL-SPECIFIC PEAK STRENGTH
    # --------------------------------------------------------

    peak_tau = (
        theoretical_failure_tau
        * trial_strength_factor
    )

    peak_tau = max(
        0.0,
        peak_tau
    )

    # --------------------------------------------------------
    # PEAK DISPLACEMENT
    # --------------------------------------------------------

    peak_fraction = rng.uniform(
        profile.peak_min,
        profile.peak_max
    )

    peak_fraction *= (
        density_factor["peak_shift"]
    )

    peak_fraction = clamp(
        peak_fraction,
        0.20,
        0.85
    )

    peak_displacement = (
        peak_fraction
        * max_displacement
    )

    # Small independent trial shift.
    # This is displacement variability, not strength variability.
    displacement_shift = rng.uniform(
        -0.025,
        0.025
    )

    peak_displacement *= (
        1.0 + displacement_shift
    )

    peak_displacement = clamp(
        peak_displacement,
        0.20 * max_displacement,
        0.85 * max_displacement
    )

    # --------------------------------------------------------
    # RESIDUAL STRENGTH
    # --------------------------------------------------------

    residual_ratio = rng.uniform(
        profile.residual_min,
        profile.residual_max
    )

    # Density modifies softening but never changes the
    # soil-specific residual range excessively.
    softening_adjustment = (
        density_factor["softening"] - 1.0
    ) * 0.08

    residual_ratio -= softening_adjustment

    residual_ratio = clamp(
        residual_ratio,
        0.65,
        1.00
    )

    residual_tau = (
        peak_tau
        * residual_ratio
    )

    # --------------------------------------------------------
    # DISPLACEMENT ARRAY
    # --------------------------------------------------------

    displacement = [
        (
            i / (n_points - 1)
        ) * max_displacement
        for i in range(n_points)
    ]

    # --------------------------------------------------------
    # TRIAL-SPECIFIC CURVE PARAMETERS
    # --------------------------------------------------------

    shape_jitter = rng.uniform(
        -0.08,
        0.08
    )

    sharpness = clamp(
        profile.peak_sharpness
        * density_factor["sharpness"]
        + shape_jitter,
        0.05,
        1.50
    )

    # Low-frequency logger drift.
    drift_amplitude = (
        profile.noise_level
        * rng.uniform(0.25, 0.60)
    )

    drift_phase = rng.uniform(
        0.0,
        2.0 * math.pi
    )

    drift_frequency = rng.uniform(
        0.55,
        1.10
    )

    # --------------------------------------------------------
    # GENERATE RESPONSE
    # --------------------------------------------------------

    shear = []

    for x in displacement:

        # ----------------------------------------------------
        # PRE-PEAK
        # ----------------------------------------------------

        if x <= peak_displacement:

            p = (
                x
                / max(peak_displacement, 1e-9)
            )

            # Three-part conceptual response:
            # seating -> mobilization -> approach to peak.

            if p <= 0.15:

                local_p = p / 0.15

                exponent = (
                    0.45
                    + 0.18 * (1.0 - sharpness)
                )

                base_factor = (
                    0.70
                    * local_p ** exponent
                )

            elif p <= 0.70:

                local_p = (
                    p - 0.15
                ) / 0.55

                exponent = (
                    0.70
                    + 0.25 * (1.0 - sharpness)
                )

                base_factor = (
                    0.70
                    + 0.24
                    * local_p ** exponent
                )

            else:

                local_p = (
                    p - 0.70
                ) / 0.30

                exponent = (
                    1.15
                    + 0.45 * sharpness
                )

                base_factor = (
                    0.94
                    + 0.06
                    * local_p ** exponent
                )

            tau = (
                peak_tau
                * base_factor
            )

        # ----------------------------------------------------
        # POST-PEAK
        # ----------------------------------------------------

        else:

            post_range = (
                max_displacement
                - peak_displacement
            )

            post_fraction = (
                (
                    x
                    - peak_displacement
                )
                / max(post_range, 1e-9)
            )

            post_fraction = clamp(
                post_fraction,
                0.0,
                1.0
            )

            softening_exponent = (
                1.15
                + 0.70 * sharpness
            )

            tau = (
                peak_tau
                - (
                    peak_tau
                    - residual_tau
                )
                * (
                    post_fraction
                    ** softening_exponent
                )
            )

        # ----------------------------------------------------
        # MACHINE / LOGGER IMPERFECTIONS
        # ----------------------------------------------------

        # Point-to-point logger noise.
        random_noise = rng.gauss(
            0.0,
            profile.noise_level
        )

        # Low-frequency drift.
        drift = (
            drift_amplitude
            * math.sin(
                drift_phase
                + drift_frequency
                * x
            )
        )

        # Occasional small local irregularity.
        local_irregularity = 0.0

        if profile.irregularity > 0:

            event_probability = (
                0.10
                + 0.08 * profile.irregularity
            )

            if rng.random() < event_probability:

                local_irregularity = (
                    rng.uniform(
                        -1.0,
                        1.0
                    )
                    * profile.irregularity
                    * 0.004
                )

        tau *= (
            1.0
            + random_noise
            + drift
            + local_irregularity
        )

        tau = max(
            0.0,
            tau
        )

        # Never permit an artificial secondary peak.
        tau = min(
            tau,
            peak_tau
        )

        shear.append(tau)

    # --------------------------------------------------------
    # FORCE PEAK OBSERVATION
    # --------------------------------------------------------

    peak_index = min(
        range(len(displacement)),
        key=lambda i: abs(
            displacement[i]
            - peak_displacement
        )
    )

    shear[peak_index] = peak_tau

    # --------------------------------------------------------
    # OBSERVED PEAK
    # --------------------------------------------------------

    observed_peak = max(shear)

    observed_peak_index = shear.index(
        observed_peak
    )

    observed_peak_displacement = (
        displacement[observed_peak_index]
    )

    # --------------------------------------------------------
    # RETURN
    # --------------------------------------------------------

    return DSTTrial(
        normal_stress=normal_stress,

        displacement=displacement,

        shear_stress=shear,

        # In a synthetic observation set, the observed peak is
        # the laboratory failure/peak stress.
        failure_shear_stress=observed_peak,

        peak_shear_stress=observed_peak,

        peak_displacement=observed_peak_displacement,

        residual_shear_stress=residual_tau,
    )


# ============================================================
# GENERATE THREE TRIALS
# ============================================================

def generate_three_trials(
    soil_type,
    density,
    c,
    phi,
    normal_stresses,
    seed=None,
    n_points=61,
):
    """
    Generate three related but independently disturbed DST trials.

    QC targets:
        c    -> target ±10% (minimum absolute tolerance 0.010)
        phi  -> target ±1.5°
        R²   -> preferably 0.970–0.990

    The R² value is NOT directly forced. Candidate sets are
    regenerated until the complete generated trials satisfy the
    engineering QC, or the best candidate is returned.

    Peak-strength variation is limited to approximately ±2–4%.
    """

    if len(normal_stresses) < 3:
        raise ValueError(
            "At least three normal stress trials are required."
        )

    master_rng = (
        random.Random()
        if seed is None
        else random.Random(int(seed))
    )

    density_class = classify_density(
        soil_type,
        density
    )

    variation = strength_variation_range(
        soil_type,
        density_class
    )

    c_tolerance = max(
        0.010,
        abs(c) * 0.10
    )

    phi_tolerance = 1.50

    preferred_r2_min = 0.970
    preferred_r2_max = 0.990

    best_trials = None
    best_score = float("inf")

    # --------------------------------------------------------
    # Controlled regeneration
    # --------------------------------------------------------

    for _attempt in range(600):

        trials = []

        # A very small shared preparation bias represents a common
        # specimen/preparation effect. Independent factors still
        # control the individual trial differences.
        shared_bias = master_rng.uniform(
            -variation * 0.15,
            variation * 0.15
        )

        for normal_stress in normal_stresses:

            independent_factor = master_rng.uniform(
                1.0 - variation,
                1.0 + variation
            )

            strength_factor = (
                independent_factor
                + shared_bias
            )

            # Keep the actual factor inside the agreed ±2–4% range.
            strength_factor = clamp(
                strength_factor,
                1.0 - variation,
                1.0 + variation
            )

            trial_seed = master_rng.randint(
                1,
                2_000_000_000
            )

            trial = generate_curve(
                soil_type=soil_type,
                density=density,
                c=c,
                phi=phi,
                normal_stress=normal_stress,
                seed=trial_seed,
                n_points=n_points,
                trial_strength_factor=strength_factor,
            )

            trials.append(trial)

        envelope = calculate_strength_envelope(
            trials
        )

        calc_c = envelope["cohesion"]
        calc_phi = envelope["phi"]
        r_squared = envelope["r_squared"]

        c_error = abs(
            calc_c - c
        )

        phi_error = abs(
            calc_phi - phi
        )

        # ----------------------------------------------------
        # R² scoring
        # ----------------------------------------------------

        if r_squared < preferred_r2_min:

            r2_penalty = (
                preferred_r2_min
                - r_squared
            ) * 12.0

        elif r_squared > preferred_r2_max:

            r2_penalty = (
                r_squared
                - preferred_r2_max
            ) * 12.0

        else:

            # Prefer the centre of the realistic range,
            # but do not force a fixed R².
            r2_penalty = (
                abs(r_squared - 0.980)
                * 0.50
            )

        score = (
            10.0
            * (
                max(
                    0.0,
                    c_error - c_tolerance
                )
                / max(c_tolerance, 1e-9)
            )
            +
            6.0
            * (
                max(
                    0.0,
                    phi_error - phi_tolerance
                )
                / phi_tolerance
            )
            +
            r2_penalty
        )

        if score < best_score:

            best_score = score
            best_trials = trials

        # ----------------------------------------------------
        # Accept candidate
        # ----------------------------------------------------

        if (
            c_error <= c_tolerance
            and phi_error <= phi_tolerance
            and preferred_r2_min <= r_squared <= preferred_r2_max
        ):

            return trials

    # --------------------------------------------------------
    # Safe fallback
    # --------------------------------------------------------

    if best_trials is None:

        raise RuntimeError(
            "Unable to generate a valid set of DST trials."
        )

    return best_trials


# ============================================================
# RESULT SUMMARY
# ============================================================

def calculate_summary(trials):
    """
    Return the complete summary expected by main.py.

    The GUI expects: 
        summary["trials"]
        summary["calculated_c"]
        summary["calculated_phi"]
        summary["r_squared"]
    """

    trial_rows = []

    # Calculate the actual regression envelope from the
    # generated peak strengths.
    envelope = calculate_strength_envelope(trials)

    for i, trial in enumerate(
        trials,
        start=1
    ):

        trial_rows.append({
            "trial": i,
            "normal_stress": trial.normal_stress,
            "failure_stress": trial.failure_shear_stress,
            "peak_stress": trial.peak_shear_stress,
            "peak_displacement": trial.peak_displacement,
            "residual_stress": trial.residual_shear_stress,
        })

    return {
        "trials": trial_rows,
        "calculated_c": envelope["cohesion"],
        "calculated_phi": envelope["phi"],
        "r_squared": envelope["r_squared"],
    }
