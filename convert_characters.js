const fs = require('fs');
const path = require('path');
const JimpObj = require('jimp');

const Jimp = JimpObj.Jimp || JimpObj;
const readImage = Jimp.read || JimpObj.read || (JimpObj.default && JimpObj.default.read);

const srcDir = 'C:\\Users\\user\\.gemini\\antigravity\\brain\\6a5b8035-61cb-4cd3-85ac-91fb293eb7e3';
const destDir = 'C:\\Users\\user\\.gemini\\antigravity\\scratch\\teunteun_expedition\\assets';

const files = [
  { src: 'new_sprout_level1_1781078549632.png', dest: 'level1.png' },
  { src: 'new_hatching_level2_1781078516434.png', dest: 'level2.png' },
  { src: 'new_level3_1781078458390.png', dest: 'level3.png' },
  { src: 'new_level4_1781078472444.png', dest: 'level4.png' }
];

async function makeBackgroundTransparent(srcFile, destFile) {
  const srcPath = path.join(srcDir, srcFile);
  const destPath = path.join(destDir, destFile);
  
  if (!fs.existsSync(srcPath)) {
    console.error(`Source file not found: ${srcPath}`);
    return;
  }
  
  console.log(`Processing ${srcFile} -> ${destFile}...`);
  const image = await readImage(srcPath);
  const width = image.bitmap.width;
  const height = image.bitmap.height;
  
  // Visited array to keep track of background pixels
  const visited = Array.from({ length: height }, () => new Uint8Array(width));
  const queue = [];
  
  // We'll treat pixels with RGB > 240 as white background
  function isWhite(x, y) {
    const idx = (y * width + x) * 4;
    const r = image.bitmap.data[idx];
    const g = image.bitmap.data[idx + 1];
    const b = image.bitmap.data[idx + 2];
    const a = image.bitmap.data[idx + 3];
    // If it's already transparent, it's background
    if (a < 50) return true;
    // If it is very close to white
    return (r > 240 && g > 240 && b > 240);
  }
  
  function enqueue(x, y) {
    if (x < 0 || x >= width || y < 0 || y >= height) return;
    if (visited[y][x]) return;
    
    if (isWhite(x, y)) {
      visited[y][x] = 1;
      queue.push([x, y]);
    }
  }
  
  // Enqueue all edge pixels
  for (let x = 0; x < width; x++) {
    enqueue(x, 0);
    enqueue(x, height - 1);
  }
  for (let y = 0; y < height; y++) {
    enqueue(0, y);
    enqueue(width - 1, y);
  }
  
  // BFS
  let head = 0;
  while (head < queue.length) {
    const [cx, cy] = queue[head++];
    enqueue(cx + 1, cy);
    enqueue(cx - 1, cy);
    enqueue(cx, cy + 1);
    enqueue(cx, cy - 1);
  }
  
  // Set all detected background pixels to fully transparent
  let count = 0;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      if (visited[y][x]) {
        const idx = (y * width + x) * 4;
        image.bitmap.data[idx + 3] = 0; // Set Alpha to 0
        count++;
      }
    }
  }
  
  await image.writeAsync(destPath);
  console.log(`Made ${count} pixels transparent in ${destFile}`);
}

async function run() {
  for (const file of files) {
    await makeBackgroundTransparent(file.src, file.dest);
  }
  console.log("All characters conversion completed!");
}

run().catch(console.error);
