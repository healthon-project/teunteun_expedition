const fs = require('fs');
const path = require('path');
const JimpObj = require('jimp');

const Jimp = JimpObj.Jimp || JimpObj;
const readImage = Jimp.read || JimpObj.read || (JimpObj.default && JimpObj.default.read);

const assetsDir = path.join(__dirname, 'assets');
const brainDir = 'C:\\Users\\user\\.gemini\\antigravity\\brain\\c4521e88-1f20-4caf-9660-29ba0828729e';

async function fixCentralHoles(filename) {
  const filePath = path.join(assetsDir, filename);
  if (!fs.existsSync(filePath)) {
    console.error(`File not found: ${filePath}`);
    return;
  }

  const image = await readImage(filePath);
  const width = image.bitmap.width;
  const height = image.bitmap.height;
  
  // Define central bounding box (characters are located in the center)
  // 1024x1024 resolution: central 180 to 840 pixels
  const xMin = Math.round(width * 0.18);
  const xMax = Math.round(width * 0.82);
  const yMin = Math.round(height * 0.18);
  const yMax = Math.round(height * 0.82);

  let filledCount = 0;

  for (let y = yMin; y <= yMax; y++) {
    for (let x = xMin; x <= xMax; x++) {
      const idx = (y * width + x) * 4;
      const alpha = image.bitmap.data[idx + 3];
      
      // If a pixel in the central area is fully or semi-transparent, fill it with white
      if (alpha < 100) {
        image.bitmap.data[idx] = 255;     // R
        image.bitmap.data[idx + 1] = 255; // G
        image.bitmap.data[idx + 2] = 255; // B
        image.bitmap.data[idx + 3] = 255; // A (opaque white)
        filledCount++;
      }
    }
  }

  if (filledCount > 0) {
    await image.writeAsync(filePath);
    console.log(`Successfully restored ${filledCount} central pixels to white in ${filename}`);
    
    // Sync to brain dir
    const brainFilePath = path.join(brainDir, filename);
    if (fs.existsSync(brainDir)) {
      fs.copyFileSync(filePath, brainFilePath);
      console.log(`Synced fixed ${filename} to brain directory.`);
    }
  } else {
    console.log(`No central transparent pixels found in ${filename}`);
  }
}

async function main() {
  try {
    // Only apply to level3.png and level4.png where holes might exist
    await fixCentralHoles('level3.png');
    await fixCentralHoles('level4.png');
    console.log("Central hole recovery finished!");
  } catch (err) {
    console.error("Error during central fix:", err);
  }
}

main();
