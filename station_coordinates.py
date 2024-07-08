"""
Copyright 2024 Wu Tingfeng <wutingfeng@outlook.com>

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
"""

import csv
import io
import re
import zipfile

import requests
import xlrd

STATION_DATA_ENDPOINT = (
    "https://datamall.lta.gov.sg/content/dam/datamall/datasets/Geospatial/"
    "Train%20Station%20Codes%20and%20Chinese%20Names.zip"
)


def to_station_code_components(station_code: str) -> tuple[str, int, str]:
    """Split station code into its components, namely, line code, station number, and station number
    suffix.

    Can be used as a key function for sorting station codes in sequential order.

    Supports station codes with alphabetical suffixes like NS3 -> NS3A -> NS4.

    Args:
        station_code (str): Station code to be split up.

    Returns:
        tuple[str, int, str]: Separated station components.
        For example ("NS", 3, "A") or ("NS", 4, "").
    """
    line_code, station_number, station_number_suffix = (
        station_code,
        0,
        "",
    )  # Default values for invalid station code.

    matcher = lambda station_code: re.match(
        "([A-Z]+)([0-9]+)([A-Z]*)", station_code
    )  # Ensure station code matches correct format.
    station_code_components_match = matcher(station_code)
    if station_code_components_match is None:
        return line_code, station_number, station_number_suffix
    matcher_groups: tuple[str, str, str] = station_code_components_match.groups("")
    line_code, station_number_str, station_number_suffix = matcher_groups
    station_number = int(station_number_str)
    return line_code, station_number, station_number_suffix


def get_stations(endpoint: str) -> list[tuple[str, str]]:
    """Download train station codes and station names.

    Args:
        endpoint (str): HTTPS address of zipped XLS file containing train station codes and names.

    Returns:
        list[tuple[str, str]]: Train stations sorted by station code in ascending order.
        For example, ("CC1", "Dhoby Ghaut"), ("NE6", "Dhoby Ghaut"), ("NS24", "Dhoby Ghaut").
    """
    with requests.Session() as session:
        res = session.get(endpoint, timeout=30)
        res.raise_for_status()
    with zipfile.ZipFile(io.BytesIO(res.content), "r") as z:
        excel_bytes = z.read(
            z.infolist()[0]
        )  # Zip file should only contain one XLS file.
        workbook = xlrd.open_workbook(file_contents=excel_bytes)
        sheet = workbook.sheet_by_index(0)

    stations: set[tuple[str, str]] = {
        (sheet.cell_value(row_idx, 0).strip(), sheet.cell_value(row_idx, 1).strip())
        for row_idx in range(1, sheet.nrows)
    }

    return sorted(
        stations,
        key=lambda station: to_station_code_components(station[0]),
    )


def get_coordinates_onemap(location_name):
    endpoint = "https://www.onemap.gov.sg/api/common/elastic/search"

    res = requests.get(
        endpoint,
        params={
            "searchVal": location_name,
            "returnGeom": "Y",
            "getAddrDetails": "Y",
            "pageNum": "1",
        },
        timeout=15,
    ).json()
    results = res.get("results", None)
    if isinstance(results, list):
        for result in results:
            if "LATITUDE" in result and "LONGITUDE" in result:
                return float(result["LATITUDE"]), float(result["LONGITUDE"])
    return None


def get_coordinates_openstreetmap(station_name):
    overpass_url = "http://overpass-api.de/api/interpreter"

    overpass_query = f"""
    [out:json];
    area["ISO3166-1"="SG"]->.searchArea;
    node[railway=station][name="{station_name}"](area.searchArea);
    out body;
    """

    response = requests.get(overpass_url, params={"data": overpass_query}, timeout=15)

    if response.status_code == 200:
        data = response.json()
    else:
        raise Exception(
            f"Error fetching data from Overpass API: {response.status_code}"
        )
    for element in data["elements"]:
        name = element.get("tags", {}).get("name", "Unnamed Station")
        lat = element["lat"]
        lon = element["lon"]
        if station_name.lower() in name.lower() and lat and lon:
            return float(lat), float(lon)
    return None


def create_kml(coordinates_file="all_stations.csv"):
    points = []
    with open(coordinates_file, "r") as f:
        csv_reader = csv.reader(f)
        next(csv_reader)
        for row in csv_reader:
            points.append((f"{row[0]} {row[1]}", row[2], row[3]))
    kml_file = coordinates_file.removesuffix(".csv") + ".kml"
    with open(kml_file, "w") as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write('<kml xmlns="http://www.opengis.net/kml/2.2">\n')
        f.write("<Document>\n")

        for name, lat, lon in points:
            f.write(f"  <Placemark>\n")
            f.write(f"    <name>{name}</name>\n")
            f.write(f"    <Point>\n")
            f.write(f"      <coordinates>{lon},{lat}</coordinates>\n")
            f.write(f"    </Point>\n")
            f.write(f"  </Placemark>\n")

        f.write("</Document>\n")
        f.write("</kml>\n")

    print(f"KML file saved as: {kml_file}")


if __name__ == "__main__":
    stations = {
        station: {
            "lat": None,
            "lon": None,
            "source": None,
            "comment": None,
        }
        for station in get_stations(STATION_DATA_ENDPOINT)
    }

    # Add future stations
    future_station_codes = set()
    with open("future_stations.csv", "r") as f:
        lines = f.readlines()
        csv_reader = csv.reader(lines)
        next(csv_reader)  # Skip header.
        for row in csv_reader:
            station_code, station_name = row[0], row[1]
            if (station_code, station_name) not in stations:
                future_station_codes.add(station_code)
                stations[(station_code, station_name)] = {
                    "lat": row[2],
                    "lon": row[3],
                    "source": row[4],
                    "comment": row[5],
                }

    for (station_code, station_name), station_details in stations.items():
        full_station_name = f"{station_code} {station_name}"
        coordinates = None
        for location_name in (
            full_station_name + " MRT",
            full_station_name + " LRT",
        ):
            try:
                coordinates = get_coordinates_onemap(location_name)
                if coordinates:
                    stations[(station_code, station_name)]["lat"] = coordinates[0]
                    stations[(station_code, station_name)]["lon"] = coordinates[1]
                    stations[(station_code, station_name)]["source"] = "onemap"
                    break
            except Exception as e:
                _ = e
        if coordinates:
            continue
        location_name = station_name
        try:
            coordinates = get_coordinates_openstreetmap(location_name)
            if coordinates:
                stations[(station_code, station_name)]["lat"] = coordinates[0]
                stations[(station_code, station_name)]["lon"] = coordinates[1]
                stations[(station_code, station_name)]["source"] = "openstreetmap"
        except Exception as e:
            _ = e

    with open("all_stations.csv", "w") as f:
        csv_writer = csv.writer(f)
        csv_writer.writerow(
            ("station_code", "station_name", "lat", "lon", "source", "comment")
        )
        csv_writer.writerows(
            sorted(
                (
                    (
                        station_code,
                        station_name,
                        *details.values(),
                    )
                    for (
                        station_code,
                        station_name,
                    ), details in stations.items()
                ),
                key=lambda x: to_station_code_components(x[0]),
            ),
        )

    with open("stations.csv", "w") as f:
        csv_writer = csv.writer(f)
        csv_writer.writerow(
            ("station_code", "station_name", "lat", "lon", "source", "comment")
        )
        csv_writer.writerows(
            sorted(
                (
                    (
                        station_code,
                        station_name,
                        *details.values(),
                    )
                    for (
                        station_code,
                        station_name,
                    ), details in stations.items()
                    if station_code not in future_station_codes
                ),
                key=lambda x: to_station_code_components(x[0]),
            ),
        )

    create_kml("all_stations.csv")
    create_kml("future_stations.csv")
    create_kml("stations.csv")
