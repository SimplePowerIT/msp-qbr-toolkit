"""Build qbr-dashboard.sppkg from qbr-dashboard.js, with no Node/SPFx toolchain.

Usage:
    python build_sppkg.py --company "Your MSP" --website https://example.com [--list "QBR Tracker"]

Upload the resulting qbr-dashboard.sppkg to your tenant app catalog, deploy it, add the
app to the site that holds the QBR Tracker list, then add the "QBR Dashboard" web part
to a page.

To ship an update: edit qbr-dashboard.js, bump VERSION below, rebuild and upload over the
old package (keep the same GUIDs so SharePoint treats it as an upgrade).
"""
import argparse, hashlib, html, json, os, zipfile

VERSION = "1.0.0"
# Fixed IDs for this solution. Generate your own once (python -c "import uuid;print(uuid.uuid4())")
# if you fork it, then keep them stable across updates.
WEBPART_ID = "5e0b7c1a-9f2d-4c3e-8a61-7d4b2f9e0c15"
PRODUCT_ID = "c3a8d2f1-6b4e-4f7a-9d20-1e5c8b7a6f34"
CFG_FEATURE = "e7f1a3b5-2c4d-4e6f-8a9b-0c1d2e3f4a5b"
CFG_ASSETS = "f8a2b4c6-3d5e-4f7a-9b0c-1d2e3f4a5b6c"
ASSETS_FEATURE = "b2da9d57-1294-5ff6-9366-9954d21ec76a"   # standard SPFx "Client Side Assets" feature ID
SP_WEBPART_BASE = ("974a7777-0990-4136-8fa6-95d80114c2e0", "1.25.0")
ADVANCED_GROUP = "5c03119e-3074-46fd-976b-c60198311f70"

X = '<?xml version="1.0" encoding="utf-8"?>'


def rels(items):
    body = "".join(f'<Relationship Type="http://schemas.microsoft.com/sharepoint/2012/app/relationships/{t}" '
                   f'Target="{target}" Id="{i}"></Relationship>' for t, target, i in items)
    return X + f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{body}</Relationships>'


def build(company, website, list_title, out):
    js = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "qbr-dashboard.js"), encoding="utf-8").read().replace("__WEBPART_ID__", f"{WEBPART_ID}_{VERSION}")
    data = js.encode("utf-8")
    jsname = f"qbr-dashboard-web-part_{hashlib.sha1(data).hexdigest()[:20]}.js"
    manifest = {
        "id": WEBPART_ID, "alias": "QbrDashboardWebPart", "componentType": "WebPart", "version": VERSION,
        "manifestVersion": 2, "requiresCustomScript": False,
        "supportedHosts": ["SharePointWebPart", "SharePointFullPage"], "supportsThemeVariants": False,
        "loaderConfig": {
            "internalModuleBaseUrls": ["HTTPS://SPCLIENTSIDEASSETLIBRARY/"], "entryModuleId": "qbr-dashboard-web-part",
            "scriptResources": {
                "qbr-dashboard-web-part": {"type": "path", "path": jsname},
                "@microsoft/sp-webpart-base": {"type": "component", "id": SP_WEBPART_BASE[0], "version": SP_WEBPART_BASE[1]}}},
        "preconfiguredEntries": [{
            "groupId": ADVANCED_GROUP, "group": {"default": "Advanced"}, "title": {"default": "QBR Dashboard"},
            "description": {"default": f"Strategic Technology Review dashboard with QBR scheduling notices, built on the {list_title} list."},
            "officeFabricIconFontName": "Calendar",
            "properties": {"listTitle": list_title, "companyName": company}}],
    }
    dev = json.dumps({"name": company, "websiteUrl": website, "privacyUrl": website, "termsOfUseUrl": website, "mpnId": "000000"})
    feat = f"feature_{WEBPART_ID}.xml"
    wp = f"{WEBPART_ID}/WebPart_{WEBPART_ID}.xml"
    files = {
        "[Content_Types].xml": X + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            + "".join(f'<Default Extension="{e}" ContentType="{c}"></Default>' for e, c in [
                ("xml", "text/xml"), ("rels", "application/vnd.openxmlformats-package.relationships+xml"),
                ("webpart", "text/xml"), ("htm", "text/html"), ("html", "text/html"), ("aspx", "text/xml"),
                ("resx", "text/xml"), ("js", "application/javascript"), ("json", "application/json"),
                ("png", "image/png"), ("jpg", "image/jpeg"), ("bmp", "image/bmp"), ("gif", "image/gif"),
                ("txt", "application/octet-stream")]) + "</Types>",
        "AppManifest.xml": X + f'<App xmlns="http://schemas.microsoft.com/sharepoint/2012/app/manifest" Name="qbr-dashboard" '
            f'ProductID="{PRODUCT_ID}" Version="{VERSION}.0" SharePointMinVersion="16.0.0.0" IsClientSideSolution="true" '
            f'SkipFeatureDeployment="false" IsDomainIsolated="false"><Properties><Title>QBR Dashboard</Title>'
            f'<DeveloperProperties>{html.escape(dev)}</DeveloperProperties></Properties></App>',
        "_rels/.rels": rels([("package-manifest", "/AppManifest.xml", "r0")]),
        "_rels/AppManifest.xml.rels": rels([("manifest-feature", f"/{feat}", "r6"), ("manifest-clientsideasset", "/ClientSideAssets.xml", "r7")]),
        feat: X + f'<Feature xmlns="http://schemas.microsoft.com/sharepoint/" Title="QBR Dashboard Feature" '
            f'Description="Adds the QBR Dashboard web part" Id="{WEBPART_ID}" Version="1.0.0.0" Scope="Web" Hidden="FALSE"></Feature>',
        f"{feat}.config.xml": X + f'<AppPartConfig xmlns="http://schemas.microsoft.com/sharepoint/2012/app/partconfiguration"><Id>{CFG_FEATURE}</Id></AppPartConfig>',
        f"_rels/{feat}.rels": rels([("partconfiguration", f"/{feat}.config.xml", "r1"), ("feature-elementmanifest", f"/{wp}", "r2")]),
        wp: X + f'<Elements xmlns="http://schemas.microsoft.com/sharepoint/"><ClientSideComponent Name="QbrDashboard" Id="{WEBPART_ID}" '
            f'ComponentManifest="{html.escape(json.dumps(manifest, separators=(",", ":")))}" Type="WebPart"></ClientSideComponent></Elements>',
        "ClientSideAssets.xml": X + f'<Feature xmlns="http://schemas.microsoft.com/sharepoint/" Title="Client Side Assets" '
            f'Description="A feature that help deploy client side component assets to SharePoint Online." Id="{ASSETS_FEATURE}" '
            f'Version="1.0.0.0" Scope="Web" Hidden="FALSE"></Feature>',
        "ClientSideAssets.xml.config.xml": X + f'<AppPartConfig xmlns="http://schemas.microsoft.com/sharepoint/2012/app/partconfiguration"><Id>{CFG_ASSETS}</Id></AppPartConfig>',
        "_rels/ClientSideAssets.xml.rels": rels([("partconfiguration", "/ClientSideAssets.xml.config.xml", "r3"), ("clientsideasset", f"/ClientSideAssets/{jsname}", "r4")]),
    }
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name, content in files.items():
            z.writestr(name, content)
        z.writestr(f"ClientSideAssets/{jsname}", data)
    print(f"Built {out} (web part {VERSION}, {jsname})")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--company", default="Your MSP", help="Shown in the scheduling email and as the app developer")
    ap.add_argument("--website", default="https://example.com")
    ap.add_argument("--list", default="QBR Tracker", help="Title of the tracker list on the site")
    ap.add_argument("--out", default="qbr-dashboard.sppkg")
    a = ap.parse_args()
    build(a.company, a.website, a.list, a.out)
