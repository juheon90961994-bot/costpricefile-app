module.exports = {
  packagerConfig: {
    asar: true,
    executableName: "Product cost Management App",
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
      description: "Product cost Management App",
      setupExe: "Product-cost-Management-App-3.1.1-Setup.exe",
      noMsi: true
    }
  }]
};
