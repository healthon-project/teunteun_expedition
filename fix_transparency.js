const fs = require('fs');
const path = require('path');
const JimpObj = require('jimp');

const Jimp = JimpObj.Jimp || JimpObj;
const readImage = Jimp.read || JimpObj.read || (JimpObj.default && JimpObj.default.read);

const assetsDir = path.join(__dirname, 'assets');
const brainDir = 'C:\\Users\\user\\.gemini\\antigravity\\brain\\c4521e88-1f20-4caf-9660-29ba0828729e';

async function fixHoles(filename) {
  const filePath = path.join(assetsDir, filename);
  if (!fs.existsSync(filePath)) {
    console.error(`File not found: ${filePath}`);
    return;
  }

  if (typeof readImage !== 'function') {
    console.error("Jimp read function is not resolved correctly.", JimpObj);
    return;
  }

  const image = await readImage(filePath);
  const width = image.bitmap.width;
  const height = image.bitmap.height;
  
  // Keep track of visited transparent outer pixels
  const visited = Array.from({ length: height }, () => new Uint8Array(width));
  const queue = [];

  // Helper to push to queue if transparent and not visited
  function enqueue(x, y) {
    if (x < 0 || x >= width || y < 0 || y >= height) return;
    if (visited[y][x]) return;
    
    const idx = (y * width + x) * 4;
    const alpha = image.bitmap.data[idx + 3];
    
    // If it is fully or mostly transparent, mark as outer background
    if (alpha < 50) { 
      visited[y][x] = 1;
      queue.push([x, y]);
    }
  }

  // Enqueue all boundary pixels
  for (let x = 0; x < width; x++) {
    enqueue(x, 0);
    enqueue(x, height - 1);
  }
  for (let y = 0; y < height; y++) {
    enqueue(0, y);
    enqueue(width - 1, y);
  }

  // BFS to find all outer transparent regions
  let head = 0;
  while (head < queue.length) {
    const [cx, cy] = queue[head++];
    enqueue(cx + 1, cy);
    enqueue(cx - 1, cy);
    enqueue(cx, cy + 1);
    enqueue(cx, cy - 1);
  }

  // Fill internal transparent holes with white
  let filledCount = 0;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const idx = (y * width + x) * 4;
      const alpha = image.bitmap.data[idx + 3];
      
      // If transparent but NOT connected to the outside boundary, it is a hole
      if (alpha < 50 && visited[y][x] === 0) {
        image.bitmap.data[idx] = 255;     // Red
        image.bitmap.data[idx + 1] = 255; // Green
        image.bitmap.data[idx + 2] = 255; // Blue
        image.bitmap.data[idx + 3] = 255; // Alpha (fully opaque white)
        filledCount++;
      }
    }
  }

  if (filledCount > 0) {
    await image.writeAsync(filePath);
    console.log(`Successfully fixed ${filledCount} transparent hole pixels in ${filename}`);
    
    // Also sync the fixed image to the brain directory for user visualization
    const brainFilePath = path.join(brainDir, filename);
    if (fs.existsSync(brainDir)) {
      fs.copyFileSync(filePath, brainFilePath);
      console.log(`Synced fixed ${filename} to brain directory.`);
    }
  } else {
    console.log(`No internal transparent holes found in ${filename}`);
  }
}

async function main() {
  try {
    await fixHoles('level1.png');
    await fixHoles('level2.png');
    await fixHoles('level3.png');
    await fixHoles('level4.png');
    console.log("All character transparency hole fixes completed!");
  } catch (err) {
    console.error("Error running fix:", err);
  }
}

main();
