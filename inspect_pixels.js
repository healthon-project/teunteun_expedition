const Jimp = require('jimp');

async function main() {
  const image = await Jimp.read('assets/level3.png');
  const width = image.bitmap.width;
  const height = image.bitmap.height;
  
  let transparentCount = 0;
  let semiTransparentCount = 0;
  let opaqueBlackCount = 0;
  
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const idx = (y * width + x) * 4;
      const r = image.bitmap.data[idx];
      const g = image.bitmap.data[idx + 1];
      const b = image.bitmap.data[idx + 2];
      const a = image.bitmap.data[idx + 3];
      
      if (a === 0) {
        transparentCount++;
      } else if (a < 255) {
        semiTransparentCount++;
      }
      
      if (r === 0 && g === 0 && b === 0 && a === 255) {
        opaqueBlackCount++;
      }
    }
  }
  
  console.log(`Image info: ${width}x${height}`);
  console.log(`Fully transparent (alpha=0): ${transparentCount}`);
  console.log(`Semi-transparent (0<alpha<255): ${semiTransparentCount}`);
  console.log(`Fully opaque black (0,0,0,255): ${opaqueBlackCount}`);
}

main();
