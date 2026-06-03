const fs = require('fs');
const path = require('path');

const srcDir = 'C:\\Users\\user\\.gemini\\antigravity\\brain\\c4521e88-1f20-4caf-9660-29ba0828729e';
const destDir = 'C:\\Users\\user\\.gemini\\antigravity\\scratch\\teunteun_expedition\\assets';

const files = [
  { src: 'cute_farm_bg_lvl1_1780463593177.png', dest: 'cute_farm_bg_lvl1.png' },
  { src: 'cute_farm_bg_lvl2_1780463606051.png', dest: 'cute_farm_bg_lvl2.png' },
  { src: 'cute_farm_bg_lvl3_1780463620350.png', dest: 'cute_farm_bg_lvl3.png' },
  { src: 'cute_farm_bg_lvl4_1780463635770.png', dest: 'cute_farm_bg_lvl4.png' }
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
