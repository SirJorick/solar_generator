# predefined_values.py
# This file defines the key component ratings and ranges used by the solar generation set calculator.
# The values below are based on real-world products and market availability.

# ---------------------------
# INVERTERS (Watts)
# ---------------------------
# Common residential and commercial inverter sizes.
INVERTER_SIZES = [
    100, 125, 150, 200, 250, 300, 350, 400, 500, 600, 750, 1000, 1500, 2000, 2500, 3000,
    4000, 5000, 6000, 8000, 10000, 15000, 20000, 25000, 30000, 40000, 50000, 60000
]

# ---------------------------
# MPPT CONTROLLERS (Amps)
# ---------------------------
# MPPT (Maximum Power Point Tracking) controllers are used to optimize the output of the PV array.
MPPT_SIZES = [
    10, 15, 20, 25, 30, 40, 50, 60, 80, 100, 120, 150, 200, 250, 300, 400, 500, 600
]

# ---------------------------
# CHARGE CONTROLLERS (Amps)
# ---------------------------
# Charge controllers regulate the power going into the battery bank from the PV array.
# They are available in PWM (Pulse Width Modulation) or MPPT types.
SCC_SIZES = [
    10, 15, 20, 25, 30, 40, 50, 60, 80, 100, 120, 150, 200, 250, 300, 400, 500
]

# ---------------------------
# DC CIRCUIT BREAKERS (Amps)
# ---------------------------
# DC breakers protect the DC side of the system.
DC_BREAKER_SIZES = [
    10, 16, 20, 25, 30, 32, 40, 50, 60, 80, 100, 120, 150, 200, 250, 300, 400, 500
]

# ---------------------------
# AC CIRCUIT BREAKERS (Amps)
# ---------------------------
# AC breakers protect the AC wiring and equipment.
BREAKER_SIZES = [
    10, 15, 16, 20, 25, 30, 32, 40, 50, 60, 70, 80, 100, 120, 150, 200, 250, 300, 400, 500
]

# ---------------------------
# CABLE SIZES (mm²)
# ---------------------------
# These are typical conductor cross-sectional areas in mm².
CABLE_SIZES = [
    1.5, 2.5, 4, 6, 10, 16, 25, 35, 50, 70, 95, 120, 150, 185, 240, 300, 400, 500, 600, 750, 900, 1200
]

# ---------------------------
# WIRE RESISTANCES (Ohms per meter)
# ---------------------------
# Typical resistances for copper conductors (values vary with temperature and conductor construction)
WIRE_RESISTANCES = {
    1.5: 0.0121,
    2.5: 0.0077,
    4:   0.0048,
    6:   0.0032,
    10:  0.0019,
    16:  0.0012,
    25:  0.00077,
    35:  0.00055,
    50:  0.00039,
    70:  0.00028,
    95:  0.00023,
    120: 0.00019,
    150: 0.00016,
    185: 0.00013,
    240: 0.00010,
    300: 0.00008,
    400: 0.00006,
    500: 0.00005,
    600: 0.00004,
    750: 0.000035,
    900: 0.000030,
    1200: 0.000024
}

# ---------------------------
# WIRE SIZE CONVERSION (mm² to AWG)
# ---------------------------
# These conversions are approximate; actual AWG values depend on insulation and conductor type.
WIRE_SIZE_CONVERSION = {
    1.5: 16, 2.5: 14, 4: 12, 6: 10, 10: 8, 16: 6, 25: 4, 35: 2, 50: 1, 70: 0, 95: "00",
    120: "000", 150: "0000"
}

# ---------------------------
# AMPACITY RATING (Amps)
# ---------------------------
# Ampacity ratings for cables (these numbers are approximate and depend on installation conditions)
AMPACITY_RATING = {
    1.5: 14, 2.5: 20, 4: 26, 6: 32, 10: 44, 16: 55, 25: 70, 35: 85, 50: 100, 70: 125, 95: 150,
    120: 170, 150: 200, 185: 230, 240: 300, 300: 350, 400: 400, 500: 500, 600: 600,
    750: 700, 900: 800, 1200: 1000
}

# ---------------------------
# SYSTEM VOLTAGES (Volts)
# ---------------------------
# Common DC system voltages for off-grid and grid-tie solar systems.
VOLTAGES = [3, 5, 7, 9, 12, 24, 36, 48, 60, 72, 84, 96, 108]

# ---------------------------
# USAGE HOURS (Hours per day)
# ---------------------------
# Typically, these represent how many hours per day an appliance runs.
USAGE_HOURS = list(range(1, 25))

# ---------------------------
# SOLAR EFFICIENCY (Percentage)
# ---------------------------
# Typical solar panel conversion efficiencies (from 15% for older panels up to 25% for high-efficiency panels)
SOLAR_EFFICIENCY = [15, 17, 20, 22, 25]

# ---------------------------
# BATTERY SIZES (Ah)
# ---------------------------
# Common battery capacities used in off-grid systems.
AH = [
    10, 20, 30, 40, 50, 60, 80, 100, 120, 150, 200, 250, 300, 400, 500, 600, 800, 1000
]

# ---------------------------
# DEPTH OF DISCHARGE (DOD) (%)
# ---------------------------
# Recommended depths-of-discharge vary by battery chemistry.
DOD = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100]

# ---------------------------
# BATTERY MANAGEMENT SYSTEM (BMS) RATING (Amps)
# ---------------------------
# BMS units protect and monitor battery packs. These values represent current handling ratings.
BMS_SIZES = [
    10, 20, 30, 40, 50, 60, 80, 100, 120, 150, 200, 250, 300, 350, 400, 450, 500, 600, 800
]

# ---------------------------
# SOLAR PANEL SIZES (Watts)
# ---------------------------
# Typical sizes for residential and small commercial panels.
PANEL_SIZES = [
    10, 20, 40, 70, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600
]

# ---------------------------
# ACTIVE BATTERY BALANCERS (Amps)
# ---------------------------
# Active battery balancers help equalize cell or battery voltages in a battery bank.
# The following list represents available balancer current ratings from low to high.
ACTIVE_BALANCER_SIZES = [
    5, 10, 15, 20, 25, 30, 40, 50, 60, 80, 100, 120, 150, 200
]

# ---------------------------
# OPTIONAL: FUSE SIZES (Amps)
# ---------------------------
# Fuse sizes are sometimes used in addition to circuit breakers.
FUSE_SIZES = [
    5, 7.5, 10, 15, 20, 25, 30, 40, 50, 60, 80, 100, 120, 150, 200
]
