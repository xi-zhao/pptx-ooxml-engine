from __future__ import annotations

import posixpath
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile

from pptx import Presentation


@dataclass
class VerifyReport:
    issues: list[str]

    @property
    def ok(self) -> bool:
        return not self.issues


def _rels_path(part_path: str) -> str:
    return posixpath.join(
        posixpath.dirname(part_path),
        "_rels",
        f"{posixpath.basename(part_path)}.rels",
    )


def _norm(base_part: str, target: str) -> str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def verify_pptx(path: str | Path) -> VerifyReport:
    issues: list[str] = []
    target = Path(path).expanduser().resolve()

    try:
        Presentation(str(target))
    except Exception as exc:
        return VerifyReport([f"python-pptx cannot open file: {exc}"])

    ns_rel = "http://schemas.openxmlformats.org/package/2006/relationships"
    ns_p = "http://schemas.openxmlformats.org/presentationml/2006/main"
    rid_attr = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

    def rel_map(archive: ZipFile, rel_path: str) -> dict[str, dict[str, str]]:
        try:
            root = ET.fromstring(archive.read(rel_path))
        except KeyError:
            issues.append(f"missing rels part: {rel_path}")
            return {}
        return {
            rel.attrib["Id"]: {"Type": rel.attrib["Type"], "Target": rel.attrib["Target"]}
            for rel in root.findall(f"{{{ns_rel}}}Relationship")
        }

    with ZipFile(target) as archive:
        pres = "ppt/presentation.xml"
        pres_rels = _rels_path(pres)
        try:
            pres_root = ET.fromstring(archive.read(pres))
        except KeyError:
            return VerifyReport(["missing ppt/presentation.xml"])
        pres_rel_map = rel_map(archive, pres_rels)

        # Registered masters.
        registered_masters: set[str] = set()
        master_lst = pres_root.find(f"{{{ns_p}}}sldMasterIdLst")
        if master_lst is not None:
            for master_id in master_lst.findall(f"{{{ns_p}}}sldMasterId"):
                rid = master_id.attrib.get(rid_attr)
                if not rid:
                    continue
                rel = pres_rel_map.get(rid)
                if rel:
                    registered_masters.add(_norm(pres, rel["Target"]))

        used_masters: set[str] = set()
        slide_lst = pres_root.find(f"{{{ns_p}}}sldIdLst")
        if slide_lst is None:
            return VerifyReport(issues)

        for slide_id in slide_lst.findall(f"{{{ns_p}}}sldId"):
            slide_rid = slide_id.attrib.get(rid_attr)
            rel = pres_rel_map.get(slide_rid or "")
            if not rel:
                issues.append(f"slide rId missing in presentation rels: {slide_rid}")
                continue
            slide_part = _norm(pres, rel["Target"])
            slide_rels = _rels_path(slide_part)
            slide_rel_map = rel_map(archive, slide_rels)
            try:
                slide_root = ET.fromstring(archive.read(slide_part))
            except KeyError:
                issues.append(f"missing slide part: {slide_part}")
                continue

            # Dangling r:id references.
            for elem in slide_root.iter():
                for attr_name, attr_value in elem.attrib.items():
                    if attr_name.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}") and attr_value:
                        if attr_value not in slide_rel_map:
                            issues.append(
                                f"dangling relationship {attr_value} in {slide_part}"
                            )

            layout_target = None
            for item in slide_rel_map.values():
                if item["Type"].endswith("/slideLayout"):
                    layout_target = item["Target"]
                    break
            if not layout_target:
                issues.append(f"missing slideLayout relation in {slide_part}")
                continue

            layout_part = _norm(slide_part, layout_target)
            layout_rels = _rels_path(layout_part)
            layout_rel_map = rel_map(archive, layout_rels)
            master_target = None
            for item in layout_rel_map.values():
                if item["Type"].endswith("/slideMaster"):
                    master_target = item["Target"]
                    break
            if not master_target:
                issues.append(f"missing slideMaster relation in {layout_part}")
                continue
            used_masters.add(_norm(layout_part, master_target))

        missing_master = sorted(used_masters - registered_masters)
        for master_part in missing_master:
            issues.append(f"used but unregistered master: {master_part}")

    return VerifyReport(issues=issues)
