const fs = require('fs');
const path = require('path');

const srcDir = 'C:\\Users\\user\\.gemini\\antigravity\\brain\\6a5b8035-61cb-4cd3-85ac-91fb293eb7e3';
const destDir = 'C:\\Users\\user\\.gemini\\antigravity\\scratch\\teunteun_expedition\\assets';

const files = [
  { src: 'perfect_half_green_land_bg_1781077799246.png', dest: 'cute_farm_bg.png' },
  { src: 'perfect_half_green_land_bg_1781077799246.png', dest: 'cute_farm_bg_lvl1.png' },
  { src: 'perfect_half_green_land_bg_1781077799246.png', dest: 'cute_farm_bg_lvl2.png' },
  { src: 'perfect_half_green_land_bg_1781077799246.png', dest: 'cute_farm_bg_lvl3.png' },
  { src: 'perfect_half_green_land_bg_1781077799246.png', dest: 'cute_farm_bg_lvl4.png' }
];

if (!fs.existsSync(destDir)) {
  fs.mkdirSync(destDir, { recursive: true });
}

files.forEach(f => {
  const srcPath = path.join(srcDir, f.src);
  const destPath = path.join(destDir, f.dest);
  if (fs.existsSync(srcPath)) {
    fs.copyFileSync(srcPath, destPath);
    console.log(`Copied: ${f.src} -> ${f.dest}`);
  } else {
    console.error(`Source not found: ${srcPath}`);
  }
});
