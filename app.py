import xml.etree.ElementTree as ET
from geopy.distance import geodesic
from io import BytesIO
import copy

def parse_kml_file_adss(kml_file, sumber):
    kml_data = kml_file.read()
    if not kml_data:
        raise ValueError("KML file is empty")

    root = ET.fromstring(kml_data)
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}

    values = {
        'tiang_new': 0,
        'tiang_existing': 0,
        'adss_12d': 0.0,
        'adss_24d': 0.0,
        'odp_8': 0,
        'odp_16': 0,
        'closure': 0,
        'otb_12': 0,
        'puas_hl': 0,
        'puas_sc': 0,
        'puas_sc_coords': [],
    }

    point_features = []

    for placemark in root.findall('.//kml:Placemark', ns):
        name_elem = placemark.find('kml:name', ns)
        desc_elem = placemark.find('kml:description', ns)

        name = name_elem.text.upper().strip() if name_elem is not None and name_elem.text else ""
        desc = desc_elem.text.upper().strip() if desc_elem is not None and desc_elem.text else ""

        if placemark.find('.//kml:Point', ns) is not None:
            point_features.append((placemark, name, desc))

            if any(tag in name for tag in ["TN", "TN7", "TIANG NEW"]):
                values['tiang_new'] += 1
            elif any(tag in name for tag in ["TE", "TIANG EXISTING"]):
                values['tiang_existing'] += 1

            if "PU-AS-HL" in desc:
                if any(kw in desc for kw in ["BEGIN", "END"]):
                    values['puas_hl'] += 1
                else:
                    values['puas_hl'] += 2
            elif any(tag in name for tag in ["TN", "TN7", "TIANG NEW", "TE", "TIANG EXISTING"]):
                values['puas_sc'] += 1
                values['puas_sc_coords'].append(placemark)

            if "ODP" in name and any(kw in name for kw in ["NEW", "BARU"]):
                if "8" in name or "ODP SOLID-PB-8 AS" in desc:
                    values['odp_8'] += 1
                elif "16" in name or "ODP SOLID-PB-16 AS" in desc:
                    values['odp_16'] += 1

            if "OTB" in name and any(kw in name for kw in ["NEW", "BARU"]):
                values['otb_12'] += 1

            if any(tag in name for tag in ["CL", "CLOSURE"]):
                values['closure'] += 1

        elif placemark.find('.//kml:LineString', ns) is not None:
            coords_elem = placemark.find('.//kml:coordinates', ns)
            if coords_elem is not None and coords_elem.text:
                coords = [
                    tuple(map(float, c.split(',')[:2]))
                    for c in coords_elem.text.split()
                    if len(c.split(',')) >= 2
                ]
                length = sum(
                    geodesic((lat1, lon1), (lat2, lon2)).meters
                    for (lon1, lat1), (lon2, lat2) in zip(coords[:-1], coords[1:])
                )

                if "ADSS-12D" in name:
                    values['adss_12d'] += length
                elif "ADSS-24D" in name:
                    values['adss_24d'] += length

    if sumber == "ODC":
        values['puas_hl'] += 1

    return values

def calculate_puas(total_odp, tiang_new, tiang_existing, tikungan):
    return max(0, (total_odp * 2) - 1 + tiang_new + tiang_existing + tikungan)

def calculate_volumes_adss(inputs):
    total_odp = inputs['odp_8'] + inputs['odp_16']

    vol_adss_12d = round(inputs.get('adss_12d', 0) * 1.02)
    vol_adss_24d = round(inputs.get('adss_24d', 0) * 1.02)

    vol_puas = calculate_puas(total_odp, inputs['tiang_new'], inputs['tiang_existing'], inputs.get('tikungan', 0))

    if inputs['sumber'] == "ODC":
        vol_os_sm_1_odc = total_odp * 2
        vol_os_sm_1_odp = 0
        vol_base_tray = 1 if vol_adss_12d > 0 else 2 if vol_adss_24d > 0 else 0
        vol_tc_02_odc = 1
        vol_dd_hdpe = 6
        vol_bc_tr = 3
    else:
        vol_os_sm_1_odc = 0
        vol_os_sm_1_odp = total_odp * 2
        vol_base_tray = 0
        vol_tc_02_odc = 0
        vol_dd_hdpe = 0
        vol_bc_tr = 0

    vol_os_sm_1 = vol_os_sm_1_odc + vol_os_sm_1_odp
    vol_pc_upc = ((total_odp - 1) // 4) + 1 if total_odp > 0 else 0
    vol_pc_apc = 18 if vol_pc_upc == 1 else vol_pc_upc * 2 if vol_pc_upc > 1 else 0
    vol_ps_1_4_odc = ((total_odp - 1) // 4) + 1 if total_odp > 0 else 0
    vol_ps_1_8_odp = 1 if inputs.get('otb_12', 0) > 0 else 0

    return [
        {"designator": "AC-OF-SM-ADSS-12D", "volume": vol_adss_12d},
        {"designator": "AC-OF-SM-ADSS-24D", "volume": vol_adss_24d},
        {"designator": "ODP Solid-PB-8 AS", "volume": inputs['odp_8']},
        {"designator": "ODP Solid-PB-16 AS", "volume": inputs['odp_16']},
        {"designator": "PU-AS-HL", "volume": inputs.get('puas_hl', 0)},
        {"designator": "PU-AS-SC", "volume": inputs.get('puas_sc', 0)},
        {"designator": "OS-SM-1-ODC", "volume": vol_os_sm_1_odc},
        {"designator": "OS-SM-1-ODP", "volume": vol_os_sm_1_odp},
        {"designator": "OS-SM-1", "volume": vol_os_sm_1},
        {"designator": "PC-UPC-652-2", "volume": vol_pc_upc},
        {"designator": "PC-APC/UPC-652-A1", "volume": vol_pc_apc},
        {"designator": "PS-1-4-ODC", "volume": vol_ps_1_4_odc},
        {"designator": "TC-02-ODC", "volume": vol_tc_02_odc},
        {"designator": "DD-HDPE-40-1", "volume": vol_dd_hdpe},
        {"designator": "BC-TR-0.6", "volume": vol_bc_tr},
        {"designator": "Base Tray ODC", "volume": vol_base_tray},
        {"designator": "SC-OF-SM-24", "volume": inputs.get('closure', 0)},
        {"designator": "TC-SM-12", "volume": inputs.get('otb_12', 0)},
        {"designator": "PS-1-8-ODP", "volume": vol_ps_1_8_odp},
        {
            "designator": "Preliminary Project HRB/Kawasan Khusus",
            "volume": 1 if inputs.get('izin') else 0,
            "izin_value": float(inputs['izin']) if inputs.get('izin') and str(inputs['izin']).replace('.', '', 1).isdigit() else 0
        }
    ]

def generate_modified_kml(original_kml_bytes, puas_sc_placemarks):
    root = ET.fromstring(original_kml_bytes)
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}

    for placemark in puas_sc_placemarks:
        desc_elem = placemark.find('kml:description', ns)
        if desc_elem is None:
            desc_elem = ET.SubElement(placemark, '{http://www.opengis.net/kml/2.2}description')
            desc_elem.text = "PU-AS-SC"
        elif "PU-AS-SC" not in desc_elem.text.upper():
            desc_elem.text = (desc_elem.text or "") + " | PU-AS-SC"

    kml_tree = ET.ElementTree(root)
    kml_bytes = BytesIO()
    kml_tree.write(kml_bytes, encoding='utf-8', xml_declaration=True)
    kml_bytes.seek(0)
    return kml_bytes
