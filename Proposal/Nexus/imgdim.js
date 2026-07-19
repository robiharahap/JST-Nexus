// minimal PNG/JPEG dimension reader (no deps)
const fs=require('fs');
function dims(file){
  const b=fs.readFileSync(file);
  if(b.slice(0,8).toString('hex')==='89504e470d0a1a0a'){ // PNG
    return {w:b.readUInt32BE(16), h:b.readUInt32BE(20), type:'png'};
  }
  if(b[0]===0xFF&&b[1]===0xD8){ // JPEG
    let o=2;
    while(o<b.length){
      if(b[o]!==0xFF){o++;continue;}
      const m=b[o+1];
      if(m>=0xC0&&m<=0xCF&&m!==0xC4&&m!==0xC8&&m!==0xCC){
        return {h:b.readUInt16BE(o+5), w:b.readUInt16BE(o+7), type:'jpg'};
      }
      o+=2+b.readUInt16BE(o+2);
    }
  }
  return null;
}
module.exports={dims};
