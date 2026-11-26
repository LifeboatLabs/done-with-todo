/*
 * Done With To Do — a tiny zip file writer.
 * Files are stored as-is (no compression), which keeps this short and dependable.
 * MIT licence.
 */
(function (root, factory) {
  var zip = factory();
  root.DWTD = root.DWTD || {};
  root.DWTD.zip = zip;
  if (typeof module === 'object' && module.exports) module.exports = zip;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  var TABLE = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    var c = 0xffffffff;
    for (var i = 0; i < bytes.length; i++) c = TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  var encoder = new TextEncoder();
  function toBytes(data) {
    if (data instanceof Uint8Array) return data;
    if (data instanceof ArrayBuffer) return new Uint8Array(data);
    return encoder.encode(String(data == null ? '' : data));
  }
  function dosTime(d) { return ((d.getHours() << 11) | (d.getMinutes() << 5) | Math.floor(d.getSeconds() / 2)) & 0xffff; }
  function dosDate(d) { return (((Math.max(1980, d.getFullYear()) - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate()) & 0xffff; }

  /**
   * files: [{ path: 'Folder/file.txt', data: string | Uint8Array | ArrayBuffer }]
   * Returns an array of Uint8Array parts; in a browser wrap them as new Blob(parts, { type: 'application/zip' }).
   */
  function makeZip(files, date) {
    date = date || new Date();
    if (files.length > 65535) throw new Error('Too many files for one zip.');
    var time = dosTime(date), day = dosDate(date);
    var parts = [], central = [], offset = 0, seen = new Set();
    files.forEach(function (f) {
      var path = String(f.path).replace(/\\/g, '/').replace(/^\/+/, '');
      if (seen.has(path.toLowerCase())) throw new Error('Duplicate file in zip: ' + path);
      seen.add(path.toLowerCase());
      var name = encoder.encode(path), data = toBytes(f.data), crc = crc32(data);

      var local = new DataView(new ArrayBuffer(30));
      local.setUint32(0, 0x04034b50, true);   // local file header
      local.setUint16(4, 20, true);           // version needed (2.0)
      local.setUint16(6, 0x0800, true);       // file names are UTF-8
      local.setUint16(8, 0, true);            // stored, no compression
      local.setUint16(10, time, true);
      local.setUint16(12, day, true);
      local.setUint32(14, crc, true);
      local.setUint32(18, data.length, true);
      local.setUint32(22, data.length, true);
      local.setUint16(26, name.length, true);
      local.setUint16(28, 0, true);
      parts.push(new Uint8Array(local.buffer), name, data);

      var entry = new DataView(new ArrayBuffer(46));
      entry.setUint32(0, 0x02014b50, true);   // central directory header
      entry.setUint16(4, 20, true);           // version made by
      entry.setUint16(6, 20, true);           // version needed
      entry.setUint16(8, 0x0800, true);
      entry.setUint16(10, 0, true);
      entry.setUint16(12, time, true);
      entry.setUint16(14, day, true);
      entry.setUint32(16, crc, true);
      entry.setUint32(20, data.length, true);
      entry.setUint32(24, data.length, true);
      entry.setUint16(28, name.length, true);
      entry.setUint16(30, 0, true);           // extra field length
      entry.setUint16(32, 0, true);           // comment length
      entry.setUint16(34, 0, true);           // disk number
      entry.setUint16(36, 0, true);           // internal attributes
      entry.setUint32(38, 0, true);           // external attributes
      entry.setUint32(42, offset, true);      // where the local header starts
      central.push(new Uint8Array(entry.buffer), name);

      offset += 30 + name.length + data.length;
      if (offset > 0xffffffff) throw new Error('Backup is too big for one zip (over 4 GB).');
    });
    var cdSize = central.reduce(function (n, p) { return n + p.length; }, 0);
    var end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);       // end of central directory
    end.setUint16(4, 0, true);
    end.setUint16(6, 0, true);
    end.setUint16(8, files.length, true);
    end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true);
    end.setUint32(16, offset, true);
    end.setUint16(20, 0, true);
    return parts.concat(central, [new Uint8Array(end.buffer)]);
  }

  return { makeZip: makeZip, crc32: crc32 };
});
