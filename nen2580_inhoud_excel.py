#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
NEN 2580 bruto/netto inhoud uit IFC met veilige fallbacks + Excel export.

- NETTO inhoud (m³): Σ IfcSpace via Qto_SpaceBaseQuantities.NetVolume / Property 'Volume',
  anders geometrisch (mits ifcopenshell.geom beschikbaar). Zie IFC Qto_SpaceBaseQuantities. 
- BRUTO inhoud (m³): IfcBuilding/IfcBuildingStorey GrossVolume, anders convex hull/bbox
  van externe elementen (wanden/slabs/dak) indien geometrie beschikbaar.

Voor officiële rapportage: NTA 2581 (type A/B) hanteren.
"""

import sys, os, argparse
from datetime import datetime
import numpy as np
import csv

# IfcOpenShell is vereist
try:
    import ifcopenshell
except Exception as e:
    print("❌ IfcOpenShell ontbreekt:", e)
    print("Installeer: pip install ifcopenshell  (of conda -c ifcopenshell)")
    sys.exit(1)

# Optioneel: geometry
HAS_GEOM = True
try:
    import ifcopenshell.geom
    settings = ifcopenshell.geom.settings()
    settings.set(settings.USE_WORLD_COORDS, True)
except Exception as e:
    HAS_GEOM = False
    print("⚠️ Geometry (ifcopenshell.geom) niet beschikbaar:", e)
    print("Het script draait door, maar geometrische fallback wordt overgeslagen.")

# Optioneel: accuratere volumes
HAS_TRIMESH = True
try:
    import trimesh
except Exception:
    HAS_TRIMESH = False

HAS_SCIPY = True
try:
    from scipy.spatial import ConvexHull
except Exception:
    HAS_SCIPY = False

def get_quantity_volume_from_obj(obj, names=("GrossVolume","NetVolume","Volume")):
    """Lees IfcQuantityVolume / Property 'Volume' uit IsDefinedBy."""
    try:
        if not hasattr(obj, "IsDefinedBy") or not obj.IsDefinedBy:
            return None
        for rel in obj.IsDefinedBy:
            rpd = getattr(rel, "RelatingPropertyDefinition", None)
            if rpd is None:
                continue
            # IfcElementQuantity
            if rpd.is_a("IfcElementQuantity"):
                for q in (rpd.Quantities or []):
                    if q.is_a("IfcQuantityVolume") and \
                       (q.Name in names or q.Name.lower() in [n.lower() for n in names]):
                        if hasattr(q, "VolumeValue") and q.VolumeValue is not None:
                            return float(q.VolumeValue)
                        try:
                            return float(q.NominalValue.wrappedValue)
                        except Exception:
                            pass
            # IfcPropertySet
            if rpd.is_a("IfcPropertySet"):
                for p in (rpd.HasProperties or []):
                    if p.Name in names or p.Name.lower() in [n.lower() for n in names]:
                        try:
                            return float(p.NominalValue.wrappedValue)
                        except Exception:
                            pass
    except Exception:
        return None
    return None

def shape_to_mesh(shape):
    try:
        geom = shape.geometry
        verts = np.array(geom.verts, dtype=np.float64).reshape(-1, 3)
        
        if hasattr(geom, 'faces'):
            faces = np.array(geom.faces, dtype=np.int64).reshape(-1, 3)
        elif hasattr(geom, 'simplices'):
            faces = np.array(geom.simplices, dtype=np.int64)
        else:
            faces = np.arange(len(verts)).reshape(-1, 3)
        
        return verts, faces
        
    except Exception as e:
        return np.empty((0,3)), np.empty((0,3))

def volume_from_mesh(verts, faces):
    if len(verts)==0 or len(faces)==0:
        return 0.0
    
    if HAS_TRIMESH:
        try:
            mesh = trimesh.Trimesh(vertices=verts, faces=faces, process=True)
            if mesh.volume > 0:
                return float(mesh.volume)
        except Exception:
            pass
    
    # signed-volume fallback
    try:
        v = 0.0
        for f in faces:
            a, b, c = verts[f[0]], verts[f[1]], verts[f[2]]
            v += np.dot(a, np.cross(b, c)) / 6.0
        v = abs(v)
        if v > 0:
            return float(v)
    except Exception:
        pass
    
    # bounding box fallback
    try:
        mins = verts.min(axis=0)
        maxs = verts.max(axis=0)
        return float(np.prod(maxs - mins))
    except Exception:
        return 0.0

def compute_net_volume_spaces(model):
    net_total = 0.0
    rows = []
    spaces = model.by_type("IfcSpace") or []
    for sp in spaces:
        name = (sp.Name or "").strip() or sp.GlobalId
        vol = get_quantity_volume_from_obj(sp, names=("NetVolume","Volume"))
        method = "Quantity(Net/Volume)" if vol is not None else "Geometry(mesh)"
        
        # DEBUG
        if vol is None:
            print(f"  Space '{name}': no quantity found. HAS_GEOM={HAS_GEOM}")
        
        if vol is None and HAS_GEOM:
            try:
                shape = ifcopenshell.geom.create_shape(settings, sp)
                v, f = shape_to_mesh(shape)
                if len(v) > 0 and len(f) > 0:
                    vol = volume_from_mesh(v, f)
                    method = "Geometry(mesh)"
                else:
                    vol = 0.0
                    method = "Unavail"
            except Exception as e:
                import traceback
                traceback.print_exc()
                vol = 0.0
                method = "Unavail"
        elif vol is None and not HAS_GEOM:
            vol = 0.0
            method = "Unavail"
        net_total += vol
        rows.append({
            "IfcType": "IfcSpace",
            "Naam": name,
            "GlobalId": sp.GlobalId,
            "Netto_m3": round(vol, 3),
            "Methode": method
        })
    return net_total, rows

def collect_building_vertices(model):
    if not HAS_GEOM:
        return np.empty((0, 3))
    verts_all = []
    def add(e):
        try:
            sh = ifcopenshell.geom.create_shape(settings, e)
            v, _ = shape_to_mesh(sh)
            if len(v):
                verts_all.append(v)
        except Exception:
            pass
    # Externe wanden conservatief meenemen
    for cls in ("IfcWallStandardCase","IfcWall"):
        for w in model.by_type(cls) or []:
            is_ext = getattr(w, "IsExternal", None)
            if is_ext in (True, None):
                add(w)
    for s in model.by_type("IfcSlab") or []:
        add(s)
    for r in model.by_type("IfcRoof") or []:
        add(r)
    if not verts_all:
        return np.empty((0,3))
    return np.vstack(verts_all)

def compute_gross_volume(model):
    # 1) quantities
    for b in model.by_type("IfcBuilding") or []:
        q = get_quantity_volume_from_obj(b, names=("GrossVolume","BrutoInhoud","Volume"))
        if q is not None:
            return float(q), "IfcBuilding:GrossVolume"
    gross_sum = 0.0; had = False
    for st in model.by_type("IfcBuildingStorey") or []:
        q = get_quantity_volume_from_obj(st, names=("GrossVolume","BrutoInhoud","Volume"))
        if q is not None:
            gross_sum += float(q); had = True
    if had and gross_sum > 0:
        return gross_sum, "Σ Storey:GrossVolume"

    # 2) geometry (convex hull) of bbox
    verts = collect_building_vertices(model)
    if len(verts):
        if HAS_SCIPY:
            try:
                hull = ConvexHull(verts)
                if hull.volume > 0:
                    return float(hull.volume), "ConvexHull(external)"
            except Exception:
                pass
        mins = verts.min(axis=0); maxs = verts.max(axis=0)
        return float(np.prod(maxs - mins)), "BoundingBox(external)"
    return 0.0, "None"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("ifc", help="Pad naar IFC")
    ap.add_argument("-o","--output", default="nen2580_inhoud_resultaten.csv")
    args = ap.parse_args()

    # Ensure output folder exists
    output_dir = os.path.dirname(args.output)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)


    if not os.path.isfile(args.ifc):
        print("❌ IFC-bestand niet gevonden:", args.ifc)
        sys.exit(1)

    # DEBUG
    print(f"HAS_GEOM: {HAS_GEOM}")
    print(f"HAS_TRIMESH: {HAS_TRIMESH}")
    print(f"HAS_SCIPY: {HAS_SCIPY}")

    # Open IFC
    model = ifcopenshell.open(args.ifc)

    # NETTO per ruimte
    net_total, rows = compute_net_volume_spaces(model)

    # BRUTO gebouw
    gross_total, gross_method = compute_gross_volume(model)

    # Write summary and spaces to CSV
    summary_file = args.output.replace('.csv', '_summary.csv')
    spaces_file = args.output.replace('.csv', '_spaces.csv')

    # Write summary
    with open(summary_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=[
            "Datum", "IFC_bestand", "Bruto_inhoud_m3", "Bruto_methode",
            "Netto_inhoud_m3 (Σ IfcSpace)", "Ruimtes_geteld"
        ])
        writer.writeheader()
        writer.writerow({
            "Datum": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "IFC_bestand": os.path.basename(args.ifc),
            "Bruto_inhoud_m3": round(gross_total, 3),
            "Bruto_methode": gross_method,
            "Netto_inhoud_m3 (Σ IfcSpace)": round(net_total, 3),
            "Ruimtes_geteld": len(rows)
        })

    # Write spaces
    if rows:
        with open(spaces_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=rows[0].keys())
            writer.writeheader()
            writer.writerows(rows)

    print(f"✅ Klaar: {summary_file} en {spaces_file}")
    print("Samenvatting: bruto/netto totaal + methode")
    print("Ruimtes: netto m³ per IfcSpace + gebruikte methode")

if __name__ == "__main__":
    main()