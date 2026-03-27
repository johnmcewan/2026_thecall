import sys
import pycdlib

def probe(iso_path):
    iso = pycdlib.PyCdlib()
    iso.open(iso_path)

    has_joliet = iso.joliet_vd is not None
    has_rr     = iso.rock_ridge
    print(f"ISO: {iso_path}")
    print(f"  Joliet:     {has_joliet}")
    print(f"  Rock Ridge: {has_rr}")
    print()

    path_type = "joliet_path" if has_joliet else "iso_path"

    sample = []
    for dir_path, dirs, files in iso.walk(**{path_type: "/"}):
        for f in files:
            sample.append(dir_path.rstrip("/") + "/" + f)
            if len(sample) >= 5:
                break
        if len(sample) >= 5:
            break

    for fp in sample:
        print(f"File: {fp}")
        for ns in ["joliet_path", "iso_path"]:
            try:
                rec = iso.get_record(**{ns: fp})
                dt  = rec.date
                print(f"  [{ns}]  type = {type(dt)}")
                for attr in dir(dt):
                    if attr.startswith("__"):
                        continue
                    val = getattr(dt, attr, None)
                    if not callable(val):
                        print(f"    .{attr} = {val!r}")
            except Exception as e:
                print(f"  [{ns}]  ERROR: {e}")
        print()

    iso.close()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python date_probe.py <path_to.iso>")
        sys.exit(1)
    probe(sys.argv[1])
