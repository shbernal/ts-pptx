# Measure LibreOffice's recomputed shape sizes (its honoring of spAutoFit / layout)
# for every named shape in a .pptx. Run with LibreOffice's bundled python.exe.
# Output: JSON { caseId: { hHmm, wHmm, xHmm, yHmm, hEmu, hPt } } to stdout.
import sys, json, uno, officehelper
from com.sun.star.beans import PropertyValue

HMM_TO_EMU = 360.0          # 1/100 mm -> EMU
HMM_TO_PT = 1.0 / 35.277778 # 1/100 mm -> points

def prop(name, val):
    p = PropertyValue(); p.Name = name; p.Value = val; return p

def main(path):
    ctx = officehelper.bootstrap()
    smgr = ctx.getServiceManager()
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    url = uno.systemPathToFileUrl(path)
    doc = desktop.loadComponentFromURL(url, "_blank", 0, (prop("Hidden", True), prop("ReadOnly", True)))
    out = {}
    try:
        pages = doc.DrawPages
        for i in range(pages.Count):
            page = pages.getByIndex(i)
            for j in range(page.Count):
                shp = page.getByIndex(j)
                name = getattr(shp, "Name", "")
                if not name:
                    continue
                sz = shp.Size; pos = shp.Position
                out[name] = {
                    "slide": i + 1,
                    "wHmm": sz.Width, "hHmm": sz.Height,
                    "xHmm": pos.X, "yHmm": pos.Y,
                    "hEmu": round(sz.Height * HMM_TO_EMU),
                    "hPt": round(sz.Height * HMM_TO_PT, 3),
                    "wPt": round(sz.Width * HMM_TO_PT, 3),
                }
    finally:
        doc.close(False)
        try:
            desktop.terminate()
        except Exception:
            pass
    print(json.dumps(out))

if __name__ == "__main__":
    main(sys.argv[1])
