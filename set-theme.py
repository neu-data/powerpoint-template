"""Write the Neudata theme (colours + fonts) into .pptx / .potx files.

Usage: python set-theme.py file1.potx [file2.pptx ...]
"""
import re
import shutil
import sys
import tempfile
import zipfile

THEME_NAME = "Neudata"
COLORS = {
    "dk1": "000000", "lt1": "FFFFFF", "dk2": "04242F", "lt2": "EAF2F5",
    "accent1": "0D7377", "accent2": "0070C0", "accent3": "055F56",
    "accent4": "0B376C", "accent5": "5FA39B", "accent6": "8A99A3",
    "hlink": "0070C0", "folHlink": "0B376C",
}
HEAD_FONT = "Century Gothic"
BODY_FONT = "Century Gothic"


def patch_theme(xml: str) -> str:
    xml = re.sub(r'(<a:theme [^>]*name=")[^"]*"', rf'\g<1>{THEME_NAME}"', xml, count=1)
    xml = re.sub(r'(<a:clrScheme name=")[^"]*"', rf'\g<1>{THEME_NAME}"', xml, count=1)
    for slot, hexval in COLORS.items():
        xml = re.sub(rf"<a:{slot}>.*?</a:{slot}>",
                     f'<a:{slot}><a:srgbClr val="{hexval}"/></a:{slot}>', xml, count=1, flags=re.S)
    xml = re.sub(r'(<a:fontScheme name=")[^"]*"', rf'\g<1>{THEME_NAME}"', xml, count=1)
    xml = re.sub(r'(<a:majorFont>\s*<a:latin typeface=")[^"]*"', rf'\g<1>{HEAD_FONT}"', xml, count=1)
    xml = re.sub(r'(<a:minorFont>\s*<a:latin typeface=")[^"]*"', rf'\g<1>{BODY_FONT}"', xml, count=1)
    return xml


def patch_file(path: str) -> None:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".zip").name
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.fullmatch(r"ppt/theme/theme\d+\.xml", item.filename):
                data = patch_theme(data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)
    print(f"theme applied: {path}")


if __name__ == "__main__":
    for f in sys.argv[1:]:
        patch_file(f)
