import pandas as pd
from timezonefinder import TimezoneFinder
import googlemaps
import pytz
from datetime import datetime

# Load Excel file
df = pd.read_excel("locations.xlsx")  # Replace with your file path

# Initialize TimezoneFinder
tf = TimezoneFinder()

# Initialize Google Maps client
gmaps = googlemaps.Client(key='GOOGLE_API_KEY')  # Replace with your Google Maps API key

# Function to get latitude and longitude using Google Maps Geocoding API
def get_lat_long(city, state):
    address = f"{city}, {state}"
    print(f"Requesting coordinates for {address}...")
    geocode_result = gmaps.geocode(address)
    if geocode_result:
        lat = geocode_result[0]["geometry"]["location"]["lat"]
        lon = geocode_result[0]["geometry"]["location"]["lng"]
        print(f"Coordinates for {address} found: (lat: {lat}, lon: {lon})")
        return lat, lon
    else:
        print(f"Error retrieving coordinates for {address}")
        return None, None

# Function to get timezone based on latitude and longitude and convert to abbreviation
def get_timezone_info(lat, lon):
    if lat is not None and lon is not None:
        timezone_name = tf.timezone_at(lat=lat, lng=lon)
        if timezone_name:
            # Convert timezone name to timezone abbreviation
            tz = pytz.timezone(timezone_name)
            now = datetime.now(tz)
            timezone_abbr = now.strftime('%Z')  # Gets the timezone abbreviation
            print(f"Timezone for ({lat}, {lon}): {timezone_name} (abbreviation: {timezone_abbr})")
            return timezone_name, timezone_abbr, lat, lon
        else:
            print("Timezone not found for given coordinates.")
            return None, None, lat, lon
    print("Coordinates not found; timezone lookup skipped.")
    return None, None, lat, lon

# Loop through each city and state in the DataFrame and add four columns for latitude, longitude, Unix timezone, and timezone abbreviation
print("Starting timezone lookup...")
df[['Unix Timezone', 'Timezone Abbr', 'Latitude', 'Longitude']] = df.apply(
    lambda row: pd.Series(get_timezone_info(*get_lat_long(row['City'], row['State']))), axis=1
)

# Save the updated DataFrame to a new Excel file
output_file = "Output.xlsx"
df.to_excel(output_file, index=False)
print(f"Timezone lookup complete. Results saved to {output_file}.")
