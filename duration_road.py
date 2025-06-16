import pandas as pd
import numpy as np
from math import radians, sin, cos, sqrt, atan2

# Data from Distance_road
cities = [
    "DME", "IKT", "KZN", "LED", "OVB", "SVO", "VKO", "VVO", "ADD", "AUH", "BAH", "CMN", "DWC", "DXB", "HRG", "IKU", "KWI", "OSS", "SHJ", "SSH", "TJU", "UTP"
]
city_names = [
    "Moscow", "Irkutsk", "Kazan", "Pulkovo", "Novosibirsk", "Moscow", "Moscow", "Vladivostok", "Addis Ababa", "Abu Dhabi", "Manama", "Casablanca", "Jebel Ali", "Dubai", "Hurghada", "Tamchy", "Kuwait City", "Osh", "Sharjah", "Sharm el-Sheikh", "Kulyab", "Rayong"
]
countries = [
    "Russia", "Russia", "Russia", "Russia", "Russia", "Russia", "Russia", "Russia", "Ethiopia", "United Arab Emirates", "Bahrain", "Morocco", "United Arab Emirates", "United Arab Emirates", "Egypt", "Kyrgyzstan", "Kuwait", "Kyrgyzstan", "United Arab Emirates", "Egypt", "Tajikistan", "Thailand"
]
latitudes = [55.7522, 52.2978, 55.7887, 59.9386, 55.0415, 55.7522, 55.7522, 43.1056, 38.79930115, 54.65110016, 50.63359833, -7.589970112, 55.161389, 55.36439896, 33.79940033, 76.713046, 47.96889877, 72.79329681, 55.51720047, 34.39500046, 69.80500031, 101.0049973]
longitudes = [37.6156, 104.296, 49.1221, 30.3141, 82.9346, 37.6156, 37.6156, 131.874, 8.977890015, 24.43300056, 26.27079964, 33.36750031, 24.896356, 25.25279999, 27.17830086, 42.58792, 29.22660065, 40.60900116, 25.32859993, 27.97730064, 37.98809814, 12.67990017]
populations = [12700000, 620000, 1300000, 5600000, 1600000, 12700000, 12700000, 600000, 392000, 1483000, 157000, 3356000, 31600, 3800000, 214000, 1427, 3400000, 450000, 1800000, 75000, 105000, 74000]

# Haversine function to calculate distance
def haversine(lat1, lon1, lat2, lon2):
    R = 6371  # Radius of Earth in kilometers
    lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
    c = 2 * atan2(sqrt(a), sqrt(1-a))
    return R * c

# Calculate distance matrix
distance_matrix = np.zeros((len(cities), len(cities)))
for i in range(len(cities)):
    for j in range(len(cities)):
        if i == j:
            distance_matrix[i][j] = 0
        else:
            distance_matrix[i][j] = haversine(latitudes[i], longitudes[i], latitudes[j], longitudes[j])

# Convert distance to time (hours) with average speed of 50 km/h
average_speed = 50  # km/h
duration_matrix = distance_matrix / average_speed

# Create DataFrame with airport codes as index and columns
df = pd.DataFrame(duration_matrix, index=cities, columns=cities)

# Add header rows with city names, countries, coordinates, and populations
df_extended = pd.DataFrame({
    "IATA": [""] + cities,
    "City": [""] + city_names,
    "Country": [""] + countries,
    "Latitude": [""] + [f"{lat}" for lat in latitudes],
    "Longitude": [""] + [f"{lon}" for lon in longitudes],
    "Population": [""] + [f"{pop}" for pop in populations]
})
df_extended = pd.concat([df_extended, df], axis=1)

# Round durations to 2 decimal places for readability
df_extended.iloc[:, 6:] = df_extended.iloc[:, 6:].round(2)

# Save to Excel
output_file = "TNDP_Russia/Duration_road_airports.xlsx"
df_extended.to_excel(output_file, index=False)

print(f"Airport duration table saved to {output_file}")