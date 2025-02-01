# Predefined sizes for each component

# Inverters (in Watts)
INVERTER_SIZES = [
    100, 125, 150, 200, 250, 300, 350, 400, 500, 600, 700, 1000, 1500, 2000, 2500, 3000, 5000, 6000, 8000,
    10000, 15000, 20000, 25000, 30000, 40000, 50000, 60000, 70000, 80000, 100000
]

# MPPT (Maximum Power Point Tracking) controllers (in Amps)
MPPT_SIZES = [
    30, 50, 100, 200, 300, 500, 1000, 2000, 3000, 5000, 6000, 8000, 10000, 15000, 20000,
    25000, 30000, 40000, 50000, 60000, 70000, 80000, 100000
]

# Charge controllers (in Amps)
SCC_SIZES = [
    30, 40, 50, 60, 80, 100, 120, 150, 200, 250, 300, 400, 500, 600, 700,
    800, 1000, 1200, 1500
]

# DC Circuit breakers (in Amps)
DC_BREAKER_SIZES = [
    16, 20, 30, 32, 40, 50, 60, 80, 100, 120, 150, 200, 250, 300, 400, 500, 600, 700,
    800, 1000, 1200, 1500
]

# AC Circuit breakers (in Amps)
BREAKER_SIZES = [
    15, 16, 20, 30, 32, 40, 50, 60, 80, 100, 120, 150, 200, 250, 300, 400, 500, 600, 700,
    800, 1000, 1200, 1500
]

# Cable sizes (in mm²)
CABLE_SIZES = [
    4, 6, 10, 16, 25, 35, 50, 70, 95, 120, 150, 185, 240, 300, 400, 500, 600,
    750, 900, 1200
]

# Wire resistances (in Ohms per meter) for each wire size
WIRE_RESISTANCES = {
    4: 0.0052,  # Ohms/m for 4 mm²
    6: 0.0033,
    10: 0.0021,
    16: 0.0013,
    25: 0.00084,
    35: 0.00059,
    50: 0.00042,
    70: 0.00030,
    95: 0.00021,
    120: 0.00016,
    150: 0.00013,
    185: 0.00011,
    240: 0.00008,
    300: 0.00006,
    400: 0.00005,
    500: 0.00004,
    600: 0.00003,
    750: 0.000025,
    900: 0.00002,
    1200: 0.000016
}

# Wire size conversion from mm² to AWG (American Wire Gauge) for compatibility with different systems
WIRE_SIZE_CONVERSION = {
    2.5: 14, 4: 12, 6: 10, 10: 8, 16: 6, 25: 4, 35: 2, 50: 1, 70: 0, 95: "00",
    120: "000", 150: "0000", 185: "0000 (quadruple 0)", 240: "250 kc-mil", 300: "300 kc-mil",
    400: "350 kc-mil", 500: "500 kc-mil", 600: "600 kc-mil", 750: "750 kc-mil", 900: "900 kc-mil"
}

# Ampacity rating (in Amps) defines the maximum current the wire can safely carry without overheating
AMPACITY_RATING = {
    2.5: 18, 4: 25, 6: 30, 10: 40, 16: 55, 25: 70, 35: 85, 50: 100, 70: 125, 95: 150,
    120: 170, 150: 200, 185: 230, 240: 300, 300: 350, 400: 400, 500: 500, 600: 600,
    750: 700, 900: 800
}

# Common system voltages (in Volts) for solar power systems
VOLTAGES = [
    3, 5, 7, 9, 12, 24, 48, 36, 60, 72, 84, 96, 108
]

# Usage hours (in hours) range from 1 to 24 hours, representing daily energy consumption
USAGE_HOURS = list(range(1, 25))

# Solar efficiency values (in percentage) represent the percentage of solar energy converted into usable electricity
SOLAR_EFFICIENCY = [
    15, 17, 22, 25
]

# Battery sizes (in Ah) define the energy storage capacity of batteries used in solar systems
AH = [
    10, 20, 30, 40, 50, 60, 80, 100, 120, 150, 200, 250, 300, 400, 500, 600
]

# Depth of Discharge (DOD) (in percentage) defines how much of the battery capacity can be used, affecting battery lifespan
DOD = [
    10, 20, 30, 40, 50, 60, 70, 80, 90, 100
]

# Predefined Battery Management System (BMS) ratings (in Amps) ensure batteries are protected and efficiently managed
BMS_SIZES = [
    10, 20, 30, 40, 50, 60, 80, 100, 120, 150, 200, 250, 300, 350, 400, 450, 500, 600
]


# Solar panel sizes (in Watts) are predefined to offer options for different energy needs
PANEL_SIZES = [
    10, 20, 40, 70, 100, 150, 200, 250, 300, 350, 400, 500
]
