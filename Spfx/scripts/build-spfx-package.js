const { spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const configDir = path.join(root, 'config');
const variantsDir = path.join(configDir, 'package-variants');
const configPath = path.join(configDir, 'config.json');
const packageSolutionPath = path.join(configDir, 'package-solution.json');
const gulpCommand = path.join(root, 'node_modules', '.bin', process.platform === 'win32' ? 'gulp.cmd' : 'gulp');

const requestedVariants = process.argv.slice(2);
const variants = requestedVariants.length > 0 ? requestedVariants : ['floating', 'frontpage'];

const originalConfig = fs.readFileSync(configPath, 'utf8');
const originalPackageSolution = fs.readFileSync(packageSolutionPath, 'utf8');

function copyVariantConfig(variant) {
  const variantDir = path.join(variantsDir, variant);
  const variantConfigPath = path.join(variantDir, 'config.json');
  const variantPackageSolutionPath = path.join(variantDir, 'package-solution.json');

  if (!fs.existsSync(variantConfigPath) || !fs.existsSync(variantPackageSolutionPath)) {
    throw new Error(`Unknown SPFx package variant: ${variant}`);
  }

  fs.copyFileSync(variantConfigPath, configPath);
  fs.copyFileSync(variantPackageSolutionPath, packageSolutionPath);
}

function runGulp(args) {
  const result = spawnSync(gulpCommand, args, {
    cwd: root,
    env: process.env,
    stdio: 'inherit',
    shell: process.platform === 'win32'
  });

  if (result.error) {
    throw result.error;
  }

  if (result.status !== 0) {
    throw new Error(`gulp ${args.join(' ')} failed with exit code ${result.status}`);
  }
}

try {
  variants.forEach((variant) => {
    console.log(`\nBuilding SPFx package variant: ${variant}`);
    copyVariantConfig(variant);
    runGulp(['clean']);
    runGulp(['bundle', '--ship']);
    runGulp(['package-solution', '--ship']);
  });
} finally {
  fs.writeFileSync(configPath, originalConfig);
  fs.writeFileSync(packageSolutionPath, originalPackageSolution);
}
