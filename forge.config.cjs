module.exports = {
  packagerConfig: {
    asar: true,
    electronZipDir: process.env.ELECTRON_ZIP_DIR,
    executableName: "Product cost Management App 4.1.0",
    ignore: filePath => {
      const normalized = filePath.replace(/\\/g, "/");
      return /(^|\/)(app|db|drizzle|examples|tests|work|worker|\.openai|\.pnpm-store|dist|out)(\/|$)/.test(normalized)
        || /(^|\/)\.git(\/|$)/.test(normalized)
        || /\/node_modules\/\.bin(\/|$)/.test(normalized)
        || /\/(package-lock\.json|yarn\.lock|pnpm-lock\.yaml)$/.test(normalized);
    }
  },
  rebuildConfig: {},
  makers: [{
    name: "@electron-forge/maker-squirrel",
    config: {
      name: "product_cost_management_app",
      authors: "ATEC Mobility",
      description: "Product cost Management App 4.1.5",
      setupExe: "Product-cost-Management-App-4.1.5-Setup.exe",
      noMsi: true
    }
  }]
};
