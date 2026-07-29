const zlib = require('zlib');
const fs = require('fs');

function makePNG(size, bg, text) {
  const w = size, h = size;
  const raw = Buffer.alloc(h * (1 + w * 4));
  for (let y = 0; y < h; y++) {
    raw[y * (w * 4 + 1)] = 0;
    for (let x = 0; x < w; x++) {
      const i = y * (w * 4 + 1) + 1 + x * 4;
      // background
      raw[i] = bg[0]; raw[i+1] = bg[1]; raw[i+2] = bg[2]; raw[i+3] = 255;
      // draw simple book icon in center area
      const cx = w/2, cy = h/2, bw = w*0.45, bh = h*0.5;
      const bx = cx - bw/2, by = cy - bh/2;
      if (x >= bx && x < bx+bw && y >= by && y < by+bh) {
        raw[i] = 255; raw[i+1] = 255; raw[i+2] = 255; raw[i+3] = 255;
      }
      // spine
      const spineW = bw * 0.08;
      if (x >= bx && x < bx+spineW && y >= by && y < by+bh) {
        raw[i] = bg[0]-30; raw[i+1] = bg[1]-30; raw[i+2] = bg[2]-30; raw[i+3] = 255;
      }
      // lines on page
      for (let li = 1; li <= 4; li++) {
        const ly = by + bh * (li / 5.5);
        if (Math.abs(y - ly) < (size > 100 ? 2 : 1) && x >= bx+spineW+4 && x < bx+bw-4) {
          raw[i] = 200; raw[i+1] = 200; raw[i+2] = 200; raw[i+3] = 255;
        }
      }
    }
  }
  const deflated = zlib.deflateSync(raw, {level: 9});

  function chunk(type, data) {
    const b = Buffer.alloc(12 + data.length);
    b.writeUInt32BE(data.length, 0);
    b.write(type, 4, 'ascii');
    data.copy(b, 8);
    const crc = crc32(Buffer.concat([Buffer.from(type, 'ascii'), data]));
    b.writeUInt32BE(crc, 8 + data.length);
    return b;
  }

  function crc32(buf) {
    let c = 0xffffffff;
    const table = crc32.table || (crc32.table = (() => {
      const t = new Uint32Array(256);
      for (let i = 0; i < 256; i++) {
        let v = i;
        for (let j = 0; j < 8; j++) v = (v & 1) ? (0xedb88320 ^ (v >>> 1)) : (v >>> 1);
        t[i] = v;
      }
      return t;
    })());
    for (const b of buf) c = table[(c ^ b) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  const sig = Buffer.from([137,80,78,71,13,10,26,10]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w,0); ihdr.writeUInt32BE(h,4);
  ihdr[8]=8; ihdr[9]=6; ihdr[10]=0; ihdr[11]=0; ihdr[12]=0;
  const idat = deflated;

  return Buffer.concat([sig, chunk('IHDR',ihdr), chunk('IDAT',idat), chunk('IEND',Buffer.alloc(0))]);
}

const bg = [26, 26, 24];
fs.writeFileSync('icon-192.png', makePNG(192, bg));
fs.writeFileSync('icon-512.png', makePNG(512, bg));
console.log('Icons generated: icon-192.png, icon-512.png');
