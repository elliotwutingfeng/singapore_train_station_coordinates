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

from collections import defaultdict
import pathlib
import csv
import io
import re
import zipfile

import requests
import xlrd
import pandas as pd
import geopandas as gpd

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

    def matcher(station_code):
        return re.match(
            "([A-Z]+)([0-9]+)([A-Z]*)", station_code
        )  # Ensure station code matches correct format.

    station_code_components_match = matcher(station_code)
    if station_code_components_match is None:
        return line_code, station_number, station_number_suffix
    matcher_groups: tuple[str, str, str] = station_code_components_match.groups("")
    line_code, station_number_str, station_number_suffix = matcher_groups
    station_number = int(station_number_str)
    return line_code, station_number, station_number_suffix


def get_operational_station_names(endpoint: str) -> list[tuple[str, str]]:
    """Download operational train station codes and station names.

    Args:
        endpoint (str): HTTPS address of zipped XLS file containing operational train station codes and names.

    Returns:
        list[tuple[str, str]]: Operational train stations sorted by station code in ascending order.
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


def create_kml(coordinates_file: str):
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
            f.write("  <Placemark>\n")
            f.write(f"    <name>{name}</name>\n")
            f.write("    <Point>\n")
            f.write(f"      <coordinates>{lon},{lat}</coordinates>\n")
            f.write("    </Point>\n")
            f.write("  </Placemark>\n")

        f.write("</Document>\n")
        f.write("</kml>\n")

    print(f"KML file saved as: {kml_file}")


if __name__ == "__main__":
    # Get list of operational stations from LTA.
    operational_stations = {
        station: {
            "lat": None,
            "lon": None,
            "source": None,
            "comment": None,
        }
        for station in get_operational_station_names(STATION_DATA_ENDPOINT)
    }

    # Get all operational stations coordinates from Master Plan
    gdf = gpd.read_file(
        pathlib.Path(__file__).parent
        / "data"
        / "MasterPlan2025RailStationLayer.geojson"
    ).to_crs(epsg=4326)
    gdf.crs = None  # Suppress warnings.
    gdf["centroid"] = gdf.geometry.centroid
    gdf["lat"] = gdf["centroid"].y
    gdf["lon"] = gdf["centroid"].x
    gdf["NAME"] = gdf["NAME"].str.upper().str.replace("INTERCHANGE", "").str.strip()
    gdf["NAME"] = gdf["NAME"].replace(
        {
            "RIVER VALLEY": "FORT CANNING",
            "JELEPANG": "JELAPANG",
            "GARDEN BY THE BAY": "GARDENS BY THE BAY",
            "ONE NORTH": "ONE-NORTH",
            "DE1": "YEW TEE VILLAGE",
            "DE2": "SUNGEI KADUT",
            "NS6": "SUNGEI KADUT",
            "SENGKANG CENTRAL": "SENGKANG",
        }
    )  # Fix known mismatches.
    gdf = gdf[
        ~gdf["NAME"].isin(["IMBIAH", "RESORTS WORLD", "BEACH"])
    ]  # Exclude Sentosa Express

    gdf = gdf[["NAME", "lat", "lon"]]

    gdf_records_dict = gdf.to_dict(orient="records")

    # TODO: Manually define separate coordinates for stations within the same interchange.
    # For now, average the coordinates.
    masterplan_stations = defaultdict(list)
    for record in gdf_records_dict:
        masterplan_stations[record["NAME"]].append((record["lat"], record["lon"]))
    # Handle stations with multiple coordinates (interchanges)
    for station_name, coordinates_list in masterplan_stations.items():
        if len(coordinates_list) == 1:
            masterplan_stations[station_name] = coordinates_list[0]
        else:
            # Average the coordinates
            avg_lat = sum(coord[0] for coord in coordinates_list) / len(
                coordinates_list
            )
            avg_lon = sum(coord[1] for coord in coordinates_list) / len(
                coordinates_list
            )
            masterplan_stations[station_name] = (avg_lat, avg_lon)

    operational_stations_with_no_coordinates = []

    for operational_station in operational_stations:
        station_code, station_name = operational_station
        station_name = station_name.upper().strip()
        _ = station_code
        if station_name in masterplan_stations:
            lat, lon = masterplan_stations[station_name]
            operational_stations[operational_station]["lat"] = float(lat)
            operational_stations[operational_station]["lon"] = float(lon)
            operational_stations[operational_station]["source"] = "ura"
        else:
            print(
                f"Warning: Missing coordinates for operational station {station_code} {station_name}"
            )
            operational_stations_with_no_coordinates.append(operational_station)

    if operational_stations_with_no_coordinates:
        raise ValueError(
            f"Missing coordinates for opened stations: {operational_stations_with_no_coordinates}"
        )

    unopened_masterplan_stations = {
        station: coordinates
        for station, coordinates in masterplan_stations.items()
        if not any(station == s[1].upper().strip() for s in operational_stations.keys())
    }

    # Update future_stations.csv
    # future_stations.csv is a manually compiled dataset
    future_stations = pd.read_csv("future_stations.csv")

    for idx, row in future_stations.iterrows():
        station_name = row["station_name"].upper().strip()
        if station_name in masterplan_stations:
            lat, lon = masterplan_stations[station_name]
            future_stations.at[idx, "lat"] = lat
            future_stations.at[idx, "lon"] = lon
            future_stations.at[idx, "source"] = "ura"
            future_stations.at[idx, "comment"] = ""
        else:
            future_stations.at[idx, "comment"] = "missing_from_masterplan"

    future_station_names = set(
        future_stations["station_name"].apply(lambda x: x.upper().strip())
    )
    extra_masterplan_stations = sorted(
        set(unopened_masterplan_stations.keys()).difference(future_station_names)
    )

    for station_name in extra_masterplan_stations:
        lat, lon = unopened_masterplan_stations[station_name]
        future_stations = pd.concat(
            [
                future_stations,
                pd.DataFrame(
                    {
                        "station_code": "",
                        "station_name": station_name.title(),
                        "lat": lat,
                        "lon": lon,
                        "source": "ura",
                        "comment": "in_masterplan_only",
                    },
                    index=[0],
                ),
            ],
            ignore_index=True,
        )

    future_stations.to_csv("future_stations.csv", index=False)

    with open("operational_stations.csv", "w") as f:
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
                    ), details in operational_stations.items()
                ),
                key=lambda x: to_station_code_components(x[0]),
            ),
        )

    operational_and_future_stations = {
        **operational_stations,
        **{
            (row["station_code"], row["station_name"]): {
                "lat": row["lat"],
                "lon": row["lon"],
                "source": row["source"],
                "comment": row["comment"],
            }
            for idx, row in future_stations.iterrows()
        },
    }

    with open("operational_and_future_stations.csv", "w") as f:
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
                    ), details in operational_and_future_stations.items()
                ),
                key=lambda x: to_station_code_components(x[0]),
            ),
        )

    with open("defunct_stations.csv", "w") as f:
        csv_writer = csv.writer(f)
        csv_writer.writerow(
            ("station_code", "station_name", "lat", "lon", "source", "comment")
        )
        csv_writer.writerows(
            sorted(
                [("BP14", "Ten Mile Junction", 1.3803369, 103.7601679, "onemap", "")],
                key=lambda x: to_station_code_components(x[0]),
            ),
        )

    create_kml("operational_and_future_stations.csv")
    create_kml("future_stations.csv")
    create_kml("operational_stations.csv")
    create_kml("defunct_stations.csv")
