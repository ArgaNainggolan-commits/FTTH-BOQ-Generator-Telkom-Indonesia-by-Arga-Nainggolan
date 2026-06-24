from lxml import etree
import math
import re


class KMLParser:

    def __init__(self, kml_file):

        self.kml_file = kml_file

        self.ns = {
            "kml":
            "http://www.opengis.net/kml/2.2"
        }

    # =====================================
    # DISTANCE
    # =====================================

    def calculate_distance(
        self,
        lat1,
        lon1,
        lat2,
        lon2
    ):

        R = 6371000

        lat1 = math.radians(lat1)
        lon1 = math.radians(lon1)

        lat2 = math.radians(lat2)
        lon2 = math.radians(lon2)

        dlat = lat2 - lat1
        dlon = lon2 - lon1

        a = (

            math.sin(dlat / 2) ** 2

            +

            math.cos(lat1)
            *
            math.cos(lat2)

            *

            math.sin(dlon / 2) ** 2

        )

        c = (

            2

            *

            math.atan2(

                math.sqrt(a),

                math.sqrt(1 - a)

            )

        )

        return R * c

    # =====================================
    # CORE
    # =====================================

    def detect_core(
        self,
        text
    ):

        if not text:
            return None

        text = text.upper()

        patterns = [

            r'(\d+)\s*CORE',

            r'KT\s*(\d+)',

            r'KU\s*ADSS\s*(\d+)',

            r'CL\s*(\d+)'

        ]

        for pattern in patterns:

            match = re.search(
                pattern,
                text
            )

            if match:

                return int(
                    match.group(1)
                )

        return None

    # =====================================
    # PORT
    # =====================================

    def detect_port(
        self,
        text
    ):

        if not text:
            return None

        match = re.search(
            r'(\d+)',
            text
        )

        if match:

            return int(
                match.group(1)
            )

        return None

    # =====================================
    # OBJECT TYPE
    # =====================================

    def detect_object_type(
        self,
        folder_name
    ):

        folder_name = (
            folder_name.upper()
        )

        if "KABEL" in folder_name or "CABLE" in folder_name:
            return "cable"

        if "ODP" in folder_name:
            return "odp"

        if "ODC" in folder_name:
            return "odc"

        if "CLOSUR" in folder_name or "CLOSURE" in folder_name:
            return "closure"

        if "HH" in folder_name or "HANDHOLE" in folder_name:
            return "handhole"

        if folder_name in ["TB", "TN"] or any(x in folder_name for x in ["TIANG BARU", "NEW POLE", "TIANG BARU"]):
            return "tb"

        if folder_name in ["TE", "TL"] or any(x in folder_name for x in ["TIANG LAMA", "TIANG EKSISTING", "EXISTING POLE"]):
            return "te"

        return "point"

    # =====================================
    # PARSE DESCRIPTION
    # =====================================

    def parse_description(self, desc_text):
        if not desc_text:
            return {}
        kv = {}
        # replace tab with space, split by lines
        lines = desc_text.replace("\t", " ").split("\n")
        for line in lines:
            if ":" in line:
                parts = line.split(":", 1)
                k = parts[0].strip().lower()
                v = parts[1].strip()
                kv[k] = v
        return kv

    # =====================================
    # POINT COORD
    # =====================================

    def extract_point_coordinate(
        self,
        placemark
    ):

        point = placemark.find(
            ".//kml:Point",
            namespaces=self.ns
        )

        if point is None:
            return None

        coord_text = point.findtext(
            ".//kml:coordinates",
            default="",
            namespaces=self.ns
        )

        if not coord_text:
            return None

        coord = coord_text.strip().split(",")

        if len(coord) < 2:
            return None

        return {

            "lon":
            float(coord[0]),

            "lat":
            float(coord[1])

        }

    # =====================================
    # LINE POINTS
    # =====================================

    def get_line_points(
        self,
        coordinate_text
    ):

        points = []

        for item in coordinate_text.strip().split():

            coord = item.split(",")

            if len(coord) < 2:
                continue

            points.append({

                "lat":
                float(coord[1]),

                "lon":
                float(coord[0])

            })

        return points

    # =====================================
    # TOTAL LINE LENGTH
    # =====================================

    def calculate_line_length(
        self,
        coordinate_text
    ):

        points = self.get_line_points(
            coordinate_text
        )

        total = 0

        for i in range(
            len(points) - 1
        ):

            total += self.calculate_distance(

                points[i]["lat"],
                points[i]["lon"],

                points[i + 1]["lat"],
                points[i + 1]["lon"]

            )

        return round(
            total,
            2
        )
    # =====================================
    # CUMULATIVE CHAINAGE
    # =====================================

    def build_chainage_table(
        self,
        line_points
    ):

        chainages = [0]

        total = 0

        for i in range(
            len(line_points) - 1
        ):

            seg_length = (
                self.calculate_distance(

                    line_points[i]["lat"],
                    line_points[i]["lon"],

                    line_points[i + 1]["lat"],
                    line_points[i + 1]["lon"]

                )
            )

            total += seg_length

            chainages.append(
                total
            )

        return chainages

    # =====================================
    # XY CONVERSION
    # =====================================

    def latlon_to_xy(
        self,
        lat,
        lon,
        ref_lat
    ):

        x = (
            lon
            *
            111320
            *
            math.cos(
                math.radians(
                    ref_lat
                )
            )
        )

        y = (
            lat
            * 111320
        )

        return x, y

    # =====================================
    # PROJECT POINT TO SEGMENT
    # =====================================

    def project_point_to_segment(
        self,
        point,
        seg_start,
        seg_end
    ):

        ref_lat = (
            point["lat"]
        )

        px, py = self.latlon_to_xy(

            point["lat"],
            point["lon"],
            ref_lat

        )

        ax, ay = self.latlon_to_xy(

            seg_start["lat"],
            seg_start["lon"],
            ref_lat

        )

        bx, by = self.latlon_to_xy(

            seg_end["lat"],
            seg_end["lon"],
            ref_lat

        )

        abx = bx - ax
        aby = by - ay

        apx = px - ax
        apy = py - ay

        ab2 = (
            abx * abx
            +
            aby * aby
        )

        if ab2 == 0:

            return {

                "t": 0,

                "distance":
                math.sqrt(
                    apx * apx
                    +
                    apy * apy
                )

            }

        t = (

            (apx * abx)
            +
            (apy * aby)

        ) / ab2

        t = max(
            0,
            min(
                1,
                t
            )
        )

        proj_x = (
            ax
            +
            t * abx
        )

        proj_y = (
            ay
            +
            t * aby
        )

        distance = math.sqrt(

            (px - proj_x) ** 2
            +
            (py - proj_y) ** 2

        )

        return {

            "t":
            t,

            "distance":
            distance

        }

    # =====================================
    # CHAINAGE ON CABLE
    # =====================================

    def calculate_pole_chainage(
        self,
        pole,
        line_points
    ):

        chainages = (
            self.build_chainage_table(
                line_points
            )
        )

        best_chainage = 0

        best_distance = (
            999999999
        )

        for i in range(
            len(line_points) - 1
        ):

            result = (
                self.project_point_to_segment(

                    pole,

                    line_points[i],

                    line_points[i + 1]

                )
            )

            if result[
                "distance"
            ] < best_distance:

                best_distance = (
                    result[
                        "distance"
                    ]
                )

                segment_length = (
                    self.calculate_distance(

                        line_points[i]["lat"],
                        line_points[i]["lon"],

                        line_points[i + 1]["lat"],
                        line_points[i + 1]["lon"]

                    )
                )

                best_chainage = (

                    chainages[i]

                    +

                    (
                        segment_length
                        *
                        result["t"]
                    )

                )

        return round(
            best_chainage,
            2
        ), best_distance

    # =====================================
    # BUILD POLE SPANS
    # =====================================

    def build_pole_spans(
        self,
        poles,
        cable_name,
        cable_core
    ):

        spans = []

        if len(poles) < 2:
            return spans

        for i in range(
            len(poles) - 1
        ):

            current = poles[i]

            next_pole = poles[
                i + 1
            ]

            span_length = (

                next_pole[
                    "chainage"
                ]

                -

                current[
                    "chainage"
                ]

            )

            if span_length < 0:
                continue

            spans.append({

                "cable":
                cable_name,

                "core":
                cable_core,

                "from_pole":
                current[
                    "pole_name"
                ],

                "to_pole":
                next_pole[
                    "pole_name"
                ],

                "length":
                round(
                    span_length,
                    2
                )

            })

        return spans
    # =====================================
    # PARSE
    # =====================================

    def parse(self):

        tree = etree.parse(
            self.kml_file
        )

        root = tree.getroot()

        raw_objects = []

        cable_spans = []

        pole_database = []

        folders = root.xpath(
            ".//kml:Folder",
            namespaces=self.ns
        )

        # =====================================
        # PASS 1
        # BUILD TB / TE DATABASE
        # =====================================

        for folder in folders:

            folder_name = folder.findtext(
                "kml:name",
                default="UNKNOWN",
                namespaces=self.ns
            )

            folder_upper = (
                folder_name
                .upper()
                .strip()
            )

            is_pole_folder = (
                folder_upper in ["TB", "TN", "TE", "TL"] or
                any(x in folder_upper for x in ["BARU", "NEW", "EKSISTING", "EXISTING", "LAMA", "TIANG", "POLE"])
            )

            if not is_pole_folder:
                continue

            placemarks = folder.xpath(
                "./kml:Placemark",
                namespaces=self.ns
            )

            if not placemarks:
                continue

            is_existing = any(x in folder_upper for x in ["TE", "TL", "EKSISTING", "EXISTING", "LAMA"])
            normalized_type = "TE" if is_existing else "TB"

            counter = 1

            for placemark in placemarks:

                point_coord = (
                    self.extract_point_coordinate(
                        placemark
                    )
                )

                if not point_coord:
                    continue

                pole_database.append({

                    "pole_name":
                    f"{normalized_type}-{counter:02d}",

                    "pole_type":
                    normalized_type,

                    "lat":
                    point_coord["lat"],

                    "lon":
                    point_coord["lon"]

                })

                counter += 1

        # =====================================
        # PASS 2
        # PROCESS OBJECTS
        # =====================================

        for folder in folders:

            folder_name = folder.findtext(
                "kml:name",
                default="UNKNOWN",
                namespaces=self.ns
            )

            placemarks = folder.xpath(
                "./kml:Placemark",
                namespaces=self.ns
            )

            for placemark in placemarks:

                name = placemark.findtext(
                    "kml:name",
                    default="",
                    namespaces=self.ns
                )

                description = placemark.findtext(
                    "kml:description",
                    default="",
                    namespaces=self.ns
                ).strip()

                desc_kv = self.parse_description(description)
                spec = desc_kv.get("specification", desc_kv.get("specification id", ""))

                core_val = self.detect_core(name)
                if not core_val and spec:
                    core_str = desc_kv.get("number of core", "")
                    if core_str.isdigit():
                        core_val = int(core_str)
                    else:
                        core_val = self.detect_core(spec)

                item = {

                    "folder":
                    folder_name,

                    "name":
                    name,

                    "description":
                    description,

                    "specification":
                    spec,

                    "object_type":
                    self.detect_object_type(
                        folder_name
                    ),

                    "core":
                    core_val,

                    "port":
                    self.detect_port(
                        name
                    ),

                    "length":
                    0,

                    "from_pole":
                    None,

                    "to_pole":
                    None

                }

                line = placemark.find(
                    ".//kml:LineString",
                    namespaces=self.ns
                )

                # =================================
                # CABLE
                # =================================

                if line is not None:

                    coord = line.findtext(
                        ".//kml:coordinates",
                        default="",
                        namespaces=self.ns
                    )

                    line_points = (
                        self.get_line_points(
                            coord
                        )
                    )

                    total_length = (
                        self.calculate_line_length(
                            coord
                        )
                    )

                    item["length"] = (
                        total_length
                    )

                    poles_on_cable = []

                    for pole in pole_database:

                        chainage, distance = (
                            self.calculate_pole_chainage(

                                pole,

                                line_points

                            )
                        )

                        if distance <= 15.0:

                            pole_copy = pole.copy()

                            pole_copy[
                                "chainage"
                            ] = chainage

                            poles_on_cable.append(
                                pole_copy
                            )

                    # ==========================
                    # SORT BY POSITION
                    # ==========================

                    poles_on_cable.sort(
                        key=lambda p:
                        p["chainage"]
                    )

                    # ==========================
                    # START END POLE
                    # ==========================

                    if len(
                        poles_on_cable
                    ) >= 2:

                        item[
                            "from_pole"
                        ] = (
                            poles_on_cable[0][
                                "pole_name"
                            ]
                        )

                        item[
                            "to_pole"
                        ] = (
                            poles_on_cable[-1][
                                "pole_name"
                            ]
                        )

                    # ==========================
                    # BUILD SPAN
                    # ==========================

                    spans = (
                        self.build_pole_spans(

                            poles_on_cable,

                            cable_name=name,

                            cable_core=item[
                                "core"
                            ]

                        )
                    )

                    cable_spans.extend(
                        spans
                    )

                raw_objects.append(
                    item
                )

        return {

            "raw_objects":
            raw_objects,

            "cable_spans":
            self.clean_spans(cable_spans, 1.0)

        }
    # =====================================
    # DEBUG
    # =====================================

    def debug(
        self,
        data,
        limit=20
    ):

        print()

        print("=" * 60)
        print("RAW OBJECTS")
        print("=" * 60)

        for item in data[
            "raw_objects"
        ][:limit]:

            print(item)

        print()

        print("=" * 60)
        print("CABLE SPANS")
        print("=" * 60)

        for item in data[
            "cable_spans"
        ][:limit]:

            print(item)

        print()

        print(
            f"TOTAL RAW OBJECTS : "
            f"{len(data['raw_objects'])}"
        )

        print(
            f"TOTAL SPANS : "
            f"{len(data['cable_spans'])}"
        )

        print()

    # =====================================
    # CLEAN SPANS
    # =====================================

    def clean_spans(
        self,
        cable_spans,
        minimum_length=1.0
    ):

        cleaned = []

        for span in cable_spans:

            length = float(
                span.get(
                    "length",
                    0
                )
            )

            if length < minimum_length:
                continue

            cleaned.append(
                span
            )

        return cleaned


# =====================================
# TEST
# =====================================

if __name__ == "__main__":

    parser = KMLParser(
        "sample.kml"
    )

    result = parser.parse()

    result["cable_spans"] = (

        parser.clean_spans(
            result[
                "cable_spans"
            ]
        )

    )

    parser.debug(
        result
    )