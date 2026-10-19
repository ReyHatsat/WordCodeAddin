// Build a ready-to-publish add-in without installing it into the local Office account.
const fs = require("node:fs");
const path = require("node:path");
const { spawnSync } = require("node:child_process");
const { exportMetadataPackage } = require("office-addin-manifest/lib/export");
const Zip = require("adm-zip");

const root = path.resolve(__dirname, "..");
const site = new URL(
  process.env.ADDIN_PRODUCTION_URL || "https://reyhatsat.github.io/WordCodeAddin/"
);
if (site.protocol !== "https:" || !site.pathname.endsWith("/") || site.search || site.hash) {
  throw new Error(
    "ADDIN_PRODUCTION_URL must be an HTTPS URL ending in / without a query or fragment."
  );
}

async function build() {
  const result = spawnSync(
    process.execPath,
    [require.resolve("webpack-cli/bin/cli.js"), "--mode", "production"],
    {
      cwd: root,
      stdio: "inherit",
      env: { ...process.env, ADDIN_PRODUCTION_URL: site.href },
    }
  );
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error("Hosted build failed.");

  const manifestPath = path.join(root, "dist/manifest.json");
  const manifest = JSON.parse(fs.readFileSync(manifestPath, "utf8"));
  if (manifest.validDomains.length !== 1 || manifest.validDomains[0] !== site.hostname) {
    throw new Error("Hosted manifest domain does not match the site.");
  }
  const page = new URL(manifest.extensions[0].runtimes[0].code.page);
  if (page.origin !== site.origin || page.pathname !== site.pathname + "taskpane.html") {
    throw new Error("Hosted task pane does not match the published path.");
  }
  if (JSON.stringify(manifest).includes("localhost"))
    throw new Error("Hosted manifest still references localhost.");
  for (const extension of manifest.extensions) {
    for (const ribbon of extension.ribbons) {
      for (const tab of ribbon.tabs) {
        for (const group of tab.groups) {
          for (const entry of [group, ...group.controls]) {
            for (const icon of entry.icons) {
              const url = new URL(icon.url);
              if (!url.href.startsWith(site.href)) throw new Error("Icon outside the hosted site.");
              if (
                !fs.existsSync(path.join(root, "dist", url.pathname.slice(site.pathname.length)))
              ) {
                throw new Error("Hosted icon file is missing.");
              }
            }
          }
        }
      }
    }
  }
  const validation = spawnSync(
    process.execPath,
    [require.resolve("office-addin-manifest/cli.js"), "validate", manifestPath],
    {
      cwd: root,
      stdio: "inherit",
    }
  );
  if (validation.error) throw validation.error;
  if (validation.status !== 0) throw new Error("Hosted manifest validation failed.");
  const packagePath = path.join(root, "dist/wordcodeaddin.zip");
  await exportMetadataPackage(packagePath, manifestPath);
  const zip = new Zip(packagePath);
  if (!zip.getEntry("en-us.json"))
    throw new Error("English translation missing from the app package.");
  if (
    JSON.parse(zip.readAsText("manifest.json")).extensions[0].runtimes[0].code.page !== page.href
  ) {
    throw new Error("App package does not contain the hosted manifest.");
  }
  fs.writeFileSync(path.join(root, "dist/.nojekyll"), "");
  console.log("Hosted panel, icons, localization, and installation package verified: " + site.href);
}

build().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
