"use strict";
var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __commonJS = (cb, mod) => function __require() {
  return mod || (0, cb[__getOwnPropNames(cb)[0]])((mod = { exports: {} }).exports, mod), mod.exports;
};
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));
var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);

// node_modules/adm-zip/util/constants.js
var require_constants = __commonJS({
  "node_modules/adm-zip/util/constants.js"(exports2, module2) {
    module2.exports = {
      /* The local file header */
      LOCHDR: 30,
      // LOC header size
      LOCSIG: 67324752,
      // "PK\003\004"
      LOCVER: 4,
      // version needed to extract
      LOCFLG: 6,
      // general purpose bit flag
      LOCHOW: 8,
      // compression method
      LOCTIM: 10,
      // modification time (2 bytes time, 2 bytes date)
      LOCCRC: 14,
      // uncompressed file crc-32 value
      LOCSIZ: 18,
      // compressed size
      LOCLEN: 22,
      // uncompressed size
      LOCNAM: 26,
      // filename length
      LOCEXT: 28,
      // extra field length
      /* The Data descriptor */
      EXTSIG: 134695760,
      // "PK\007\008"
      EXTHDR: 16,
      // EXT header size
      EXTCRC: 4,
      // uncompressed file crc-32 value
      EXTSIZ: 8,
      // compressed size
      EXTLEN: 12,
      // uncompressed size
      /* The central directory file header */
      CENHDR: 46,
      // CEN header size
      CENSIG: 33639248,
      // "PK\001\002"
      CENVEM: 4,
      // version made by
      CENVER: 6,
      // version needed to extract
      CENFLG: 8,
      // encrypt, decrypt flags
      CENHOW: 10,
      // compression method
      CENTIM: 12,
      // modification time (2 bytes time, 2 bytes date)
      CENCRC: 16,
      // uncompressed file crc-32 value
      CENSIZ: 20,
      // compressed size
      CENLEN: 24,
      // uncompressed size
      CENNAM: 28,
      // filename length
      CENEXT: 30,
      // extra field length
      CENCOM: 32,
      // file comment length
      CENDSK: 34,
      // volume number start
      CENATT: 36,
      // internal file attributes
      CENATX: 38,
      // external file attributes (host system dependent)
      CENOFF: 42,
      // LOC header offset
      /* The entries in the end of central directory */
      ENDHDR: 22,
      // END header size
      ENDSIG: 101010256,
      // "PK\005\006"
      ENDSUB: 8,
      // number of entries on this disk
      ENDTOT: 10,
      // total number of entries
      ENDSIZ: 12,
      // central directory size in bytes
      ENDOFF: 16,
      // offset of first CEN header
      ENDCOM: 20,
      // zip file comment length
      END64HDR: 20,
      // zip64 END header size
      END64SIG: 117853008,
      // zip64 Locator signature, "PK\006\007"
      END64START: 4,
      // number of the disk with the start of the zip64
      END64OFF: 8,
      // relative offset of the zip64 end of central directory
      END64NUMDISKS: 16,
      // total number of disks
      ZIP64SIG: 101075792,
      // zip64 signature, "PK\006\006"
      ZIP64HDR: 56,
      // zip64 record minimum size
      ZIP64LEAD: 12,
      // leading bytes at the start of the record, not counted by the value stored in ZIP64SIZE
      ZIP64SIZE: 4,
      // zip64 size of the central directory record
      ZIP64VEM: 12,
      // zip64 version made by
      ZIP64VER: 14,
      // zip64 version needed to extract
      ZIP64DSK: 16,
      // zip64 number of this disk
      ZIP64DSKDIR: 20,
      // number of the disk with the start of the record directory
      ZIP64SUB: 24,
      // number of entries on this disk
      ZIP64TOT: 32,
      // total number of entries
      ZIP64SIZB: 40,
      // zip64 central directory size in bytes
      ZIP64OFF: 48,
      // offset of start of central directory with respect to the starting disk number
      ZIP64EXTRA: 56,
      // extensible data sector
      /* Compression methods */
      STORED: 0,
      // no compression
      SHRUNK: 1,
      // shrunk
      REDUCED1: 2,
      // reduced with compression factor 1
      REDUCED2: 3,
      // reduced with compression factor 2
      REDUCED3: 4,
      // reduced with compression factor 3
      REDUCED4: 5,
      // reduced with compression factor 4
      IMPLODED: 6,
      // imploded
      // 7 reserved for Tokenizing compression algorithm
      DEFLATED: 8,
      // deflated
      ENHANCED_DEFLATED: 9,
      // enhanced deflated
      PKWARE: 10,
      // PKWare DCL imploded
      // 11 reserved by PKWARE
      BZIP2: 12,
      //  compressed using BZIP2
      // 13 reserved by PKWARE
      LZMA: 14,
      // LZMA
      // 15-17 reserved by PKWARE
      IBM_TERSE: 18,
      // compressed using IBM TERSE
      IBM_LZ77: 19,
      // IBM LZ77 z
      AES_ENCRYPT: 99,
      // WinZIP AES encryption method
      /* General purpose bit flag */
      // values can obtained with expression 2**bitnr
      FLG_ENC: 1,
      // Bit 0: encrypted file
      FLG_COMP1: 2,
      // Bit 1, compression option
      FLG_COMP2: 4,
      // Bit 2, compression option
      FLG_DESC: 8,
      // Bit 3, data descriptor
      FLG_ENH: 16,
      // Bit 4, enhanced deflating
      FLG_PATCH: 32,
      // Bit 5, indicates that the file is compressed patched data.
      FLG_STR: 64,
      // Bit 6, strong encryption (patented)
      // Bits 7-10: Currently unused.
      FLG_EFS: 2048,
      // Bit 11: Language encoding flag (EFS)
      // Bit 12: Reserved by PKWARE for enhanced compression.
      // Bit 13: encrypted the Central Directory (patented).
      // Bits 14-15: Reserved by PKWARE.
      FLG_MSK: 4096,
      // mask header values
      /* Load type */
      FILE: 2,
      BUFFER: 1,
      NONE: 0,
      /* 4.5 Extensible data fields */
      EF_ID: 0,
      EF_SIZE: 2,
      /* Header IDs */
      ID_ZIP64: 1,
      ID_AVINFO: 7,
      ID_PFS: 8,
      ID_OS2: 9,
      ID_NTFS: 10,
      ID_OPENVMS: 12,
      ID_UNIX: 13,
      ID_FORK: 14,
      ID_PATCH: 15,
      ID_X509_PKCS7: 20,
      ID_X509_CERTID_F: 21,
      ID_X509_CERTID_C: 22,
      ID_STRONGENC: 23,
      ID_RECORD_MGT: 24,
      ID_X509_PKCS7_RL: 25,
      ID_IBM1: 101,
      ID_IBM2: 102,
      ID_POSZIP: 18064,
      EF_ZIP64_OR_32: 4294967295,
      EF_ZIP64_OR_16: 65535,
      EF_ZIP64_SUNCOMP: 0,
      EF_ZIP64_SCOMP: 8,
      EF_ZIP64_RHO: 16,
      EF_ZIP64_DSN: 24
    };
  }
});

// node_modules/adm-zip/util/errors.js
var require_errors = __commonJS({
  "node_modules/adm-zip/util/errors.js"(exports2) {
    var errors = {
      /* Header error messages */
      INVALID_LOC: "Invalid LOC header (bad signature)",
      INVALID_CEN: "Invalid CEN header (bad signature)",
      INVALID_END: "Invalid END header (bad signature)",
      /* Descriptor */
      DESCRIPTOR_NOT_EXIST: "No descriptor present",
      DESCRIPTOR_UNKNOWN: "Unknown descriptor format",
      DESCRIPTOR_FAULTY: "Descriptor data is malformed",
      /* ZipEntry error messages*/
      NO_DATA: "Nothing to decompress",
      BAD_CRC: "CRC32 checksum failed {0}",
      FILE_IN_THE_WAY: "There is a file in the way: {0}",
      UNKNOWN_METHOD: "Invalid/unsupported compression method",
      /* Inflater error messages */
      AVAIL_DATA: "inflate::Available inflate data did not terminate",
      INVALID_DISTANCE: "inflate::Invalid literal/length or distance code in fixed or dynamic block",
      TO_MANY_CODES: "inflate::Dynamic block code description: too many length or distance codes",
      INVALID_REPEAT_LEN: "inflate::Dynamic block code description: repeat more than specified lengths",
      INVALID_REPEAT_FIRST: "inflate::Dynamic block code description: repeat lengths with no first length",
      INCOMPLETE_CODES: "inflate::Dynamic block code description: code lengths codes incomplete",
      INVALID_DYN_DISTANCE: "inflate::Dynamic block code description: invalid distance code lengths",
      INVALID_CODES_LEN: "inflate::Dynamic block code description: invalid literal/length code lengths",
      INVALID_STORE_BLOCK: "inflate::Stored block length did not match one's complement",
      INVALID_BLOCK_TYPE: "inflate::Invalid block type (type == 3)",
      /* ADM-ZIP error messages */
      CANT_EXTRACT_FILE: "Could not extract the file",
      CANT_OVERRIDE: "Target file already exists",
      DISK_ENTRY_TOO_LARGE: "Number of disk entries is too large",
      NO_ZIP: "No zip file was loaded",
      NO_ENTRY: "Entry doesn't exist",
      DIRECTORY_CONTENT_ERROR: "A directory cannot have content",
      FILE_NOT_FOUND: 'File not found: "{0}"',
      NOT_IMPLEMENTED: "Not implemented",
      INVALID_FILENAME: "Invalid filename",
      INVALID_FORMAT: "Invalid or unsupported zip format. No END header found",
      INVALID_PASS_PARAM: "Incompatible password parameter",
      WRONG_PASSWORD: "Wrong Password",
      /* ADM-ZIP */
      COMMENT_TOO_LONG: "Comment is too long",
      // Comment can be max 65535 bytes long (NOTE: some non-US characters may take more space)
      EXTRA_FIELD_PARSE_ERROR: "Extra field parsing error"
    };
    function E(message) {
      return function(...args) {
        if (args.length) {
          message = message.replace(/\{(\d)\}/g, (_, n) => args[n] || "");
        }
        return new Error("ADM-ZIP: " + message);
      };
    }
    for (const msg of Object.keys(errors)) {
      exports2[msg] = E(errors[msg]);
    }
  }
});

// node_modules/adm-zip/util/utils.js
var require_utils = __commonJS({
  "node_modules/adm-zip/util/utils.js"(exports2, module2) {
    var fsystem = require("fs");
    var pth = require("path");
    var Constants = require_constants();
    var Errors = require_errors();
    var isWin = typeof process === "object" && "win32" === process.platform;
    var is_Obj = (obj) => typeof obj === "object" && obj !== null;
    var crcTable = new Uint32Array(256).map((t, c) => {
      for (let k = 0; k < 8; k++) {
        if ((c & 1) !== 0) {
          c = 3988292384 ^ c >>> 1;
        } else {
          c >>>= 1;
        }
      }
      return c >>> 0;
    });
    function Utils(opts) {
      this.sep = pth.sep;
      this.fs = fsystem;
      if (is_Obj(opts)) {
        if (is_Obj(opts.fs) && typeof opts.fs.statSync === "function") {
          this.fs = opts.fs;
        }
      }
    }
    module2.exports = Utils;
    Utils.prototype.makeDir = function(folder) {
      const self = this;
      function mkdirSync2(fpath) {
        let resolvedPath = fpath.split(self.sep)[0];
        fpath.split(self.sep).forEach(function(name) {
          if (!name || name.substr(-1, 1) === ":") return;
          resolvedPath += self.sep + name;
          var stat;
          try {
            stat = self.fs.statSync(resolvedPath);
          } catch (e) {
            self.fs.mkdirSync(resolvedPath);
          }
          if (stat && stat.isFile()) throw Errors.FILE_IN_THE_WAY(`"${resolvedPath}"`);
        });
      }
      mkdirSync2(folder);
    };
    Utils.prototype.writeFileTo = function(path2, content, overwrite, attr) {
      const self = this;
      if (self.fs.existsSync(path2)) {
        if (!overwrite) return false;
        var stat = self.fs.statSync(path2);
        if (stat.isDirectory()) {
          return false;
        }
      }
      var folder = pth.dirname(path2);
      if (!self.fs.existsSync(folder)) {
        self.makeDir(folder);
      }
      var fd;
      try {
        fd = self.fs.openSync(path2, "w", 438);
      } catch (e) {
        self.fs.chmodSync(path2, 438);
        fd = self.fs.openSync(path2, "w", 438);
      }
      if (fd) {
        try {
          self.fs.writeSync(fd, content, 0, content.length, 0);
        } finally {
          self.fs.closeSync(fd);
        }
      }
      self.fs.chmodSync(path2, attr || 438);
      return true;
    };
    Utils.prototype.writeFileToAsync = function(path2, content, overwrite, attr, callback) {
      if (typeof attr === "function") {
        callback = attr;
        attr = void 0;
      }
      const self = this;
      self.fs.exists(path2, function(exist) {
        if (exist && !overwrite) return callback(false);
        self.fs.stat(path2, function(err, stat) {
          if (exist && stat.isDirectory()) {
            return callback(false);
          }
          var folder = pth.dirname(path2);
          self.fs.exists(folder, function(exists) {
            if (!exists) self.makeDir(folder);
            self.fs.open(path2, "w", 438, function(err2, fd) {
              if (err2) {
                self.fs.chmod(path2, 438, function() {
                  self.fs.open(path2, "w", 438, function(err3, fd2) {
                    self.fs.write(fd2, content, 0, content.length, 0, function() {
                      self.fs.close(fd2, function() {
                        self.fs.chmod(path2, attr || 438, function() {
                          callback(true);
                        });
                      });
                    });
                  });
                });
              } else if (fd) {
                self.fs.write(fd, content, 0, content.length, 0, function() {
                  self.fs.close(fd, function() {
                    self.fs.chmod(path2, attr || 438, function() {
                      callback(true);
                    });
                  });
                });
              } else {
                self.fs.chmod(path2, attr || 438, function() {
                  callback(true);
                });
              }
            });
          });
        });
      });
    };
    Utils.prototype.findFiles = function(path2) {
      const self = this;
      function findSync(dir, pattern, recursive) {
        if (typeof pattern === "boolean") {
          recursive = pattern;
          pattern = void 0;
        }
        let files = [];
        self.fs.readdirSync(dir).forEach(function(file) {
          const path3 = pth.join(dir, file);
          const stat = self.fs.statSync(path3);
          if (!pattern || pattern.test(path3)) {
            files.push(pth.normalize(path3) + (stat.isDirectory() ? self.sep : ""));
          }
          if (stat.isDirectory() && recursive) files = files.concat(findSync(path3, pattern, recursive));
        });
        return files;
      }
      return findSync(path2, void 0, true);
    };
    Utils.prototype.findFilesAsync = function(dir, cb) {
      const self = this;
      let results = [];
      self.fs.readdir(dir, function(err, list) {
        if (err) return cb(err);
        let list_length = list.length;
        if (!list_length) return cb(null, results);
        list.forEach(function(file) {
          file = pth.join(dir, file);
          self.fs.stat(file, function(err2, stat) {
            if (err2) return cb(err2);
            if (stat) {
              results.push(pth.normalize(file) + (stat.isDirectory() ? self.sep : ""));
              if (stat.isDirectory()) {
                self.findFilesAsync(file, function(err3, res) {
                  if (err3) return cb(err3);
                  results = results.concat(res);
                  if (!--list_length) cb(null, results);
                });
              } else {
                if (!--list_length) cb(null, results);
              }
            }
          });
        });
      });
    };
    Utils.prototype.getAttributes = function() {
    };
    Utils.prototype.setAttributes = function() {
    };
    Utils.crc32update = function(crc, byte) {
      return crcTable[(crc ^ byte) & 255] ^ crc >>> 8;
    };
    Utils.crc32 = function(buf) {
      if (typeof buf === "string") {
        buf = Buffer.from(buf, "utf8");
      }
      let len = buf.length;
      let crc = ~0;
      for (let off = 0; off < len; ) crc = Utils.crc32update(crc, buf[off++]);
      return ~crc >>> 0;
    };
    Utils.methodToString = function(method) {
      switch (method) {
        case Constants.STORED:
          return "STORED (" + method + ")";
        case Constants.DEFLATED:
          return "DEFLATED (" + method + ")";
        default:
          return "UNSUPPORTED (" + method + ")";
      }
    };
    Utils.canonical = function(path2) {
      if (!path2) return "";
      const safeSuffix = pth.posix.normalize("/" + path2.split("\\").join("/"));
      return pth.join(".", safeSuffix);
    };
    Utils.zipnamefix = function(path2) {
      if (!path2) return "";
      const safeSuffix = pth.posix.normalize("/" + path2.split("\\").join("/"));
      return pth.posix.join(".", safeSuffix);
    };
    Utils.findLast = function(arr, callback) {
      if (!Array.isArray(arr)) throw new TypeError("arr is not array");
      const len = arr.length >>> 0;
      for (let i = len - 1; i >= 0; i--) {
        if (callback(arr[i], i, arr)) {
          return arr[i];
        }
      }
      return void 0;
    };
    Utils.sanitize = function(prefix, name) {
      prefix = pth.resolve(pth.normalize(prefix));
      var parts = name.split("/");
      for (var i = 0, l = parts.length; i < l; i++) {
        var path2 = pth.normalize(pth.join(prefix, parts.slice(i, l).join(pth.sep)));
        if (path2.indexOf(prefix) === 0) {
          return path2;
        }
      }
      return pth.normalize(pth.join(prefix, pth.basename(name)));
    };
    Utils.toBuffer = function toBuffer(input, encoder) {
      if (Buffer.isBuffer(input)) {
        return input;
      } else if (input instanceof Uint8Array) {
        return Buffer.from(input);
      } else {
        return typeof input === "string" ? encoder(input) : Buffer.alloc(0);
      }
    };
    Utils.readBigUInt64LE = function(buffer, index) {
      var slice = Buffer.from(buffer.slice(index, index + 8));
      slice.swap64();
      return parseInt(`0x${slice.toString("hex")}`);
    };
    Utils.fromDOS2Date = function(val) {
      return new Date((val >> 25 & 127) + 1980, Math.max((val >> 21 & 15) - 1, 0), Math.max(val >> 16 & 31, 1), val >> 11 & 31, val >> 5 & 63, (val & 31) << 1);
    };
    Utils.fromDate2DOS = function(val) {
      let date = 0;
      let time = 0;
      if (val.getFullYear() > 1979) {
        date = (val.getFullYear() - 1980 & 127) << 9 | val.getMonth() + 1 << 5 | val.getDate();
        time = val.getHours() << 11 | val.getMinutes() << 5 | val.getSeconds() >> 1;
      }
      return date << 16 | time;
    };
    Utils.isWin = isWin;
    Utils.crcTable = crcTable;
  }
});

// node_modules/adm-zip/util/fattr.js
var require_fattr = __commonJS({
  "node_modules/adm-zip/util/fattr.js"(exports2, module2) {
    var pth = require("path");
    module2.exports = function(path2, { fs: fs2 }) {
      var _path = path2 || "", _obj = newAttr(), _stat = null;
      function newAttr() {
        return {
          directory: false,
          readonly: false,
          hidden: false,
          executable: false,
          mtime: 0,
          atime: 0
        };
      }
      if (_path && fs2.existsSync(_path)) {
        _stat = fs2.statSync(_path);
        _obj.directory = _stat.isDirectory();
        _obj.mtime = _stat.mtime;
        _obj.atime = _stat.atime;
        _obj.executable = (73 & _stat.mode) !== 0;
        _obj.readonly = (128 & _stat.mode) === 0;
        _obj.hidden = pth.basename(_path)[0] === ".";
      } else {
        console.warn("Invalid path: " + _path);
      }
      return {
        get directory() {
          return _obj.directory;
        },
        get readOnly() {
          return _obj.readonly;
        },
        get hidden() {
          return _obj.hidden;
        },
        get mtime() {
          return _obj.mtime;
        },
        get atime() {
          return _obj.atime;
        },
        get executable() {
          return _obj.executable;
        },
        decodeAttributes: function() {
        },
        encodeAttributes: function() {
        },
        toJSON: function() {
          return {
            path: _path,
            isDirectory: _obj.directory,
            isReadOnly: _obj.readonly,
            isHidden: _obj.hidden,
            isExecutable: _obj.executable,
            mTime: _obj.mtime,
            aTime: _obj.atime
          };
        },
        toString: function() {
          return JSON.stringify(this.toJSON(), null, "	");
        }
      };
    };
  }
});

// node_modules/adm-zip/util/decoder.js
var require_decoder = __commonJS({
  "node_modules/adm-zip/util/decoder.js"(exports2, module2) {
    module2.exports = {
      efs: true,
      encode: (data) => Buffer.from(data, "utf8"),
      decode: (data) => data.toString("utf8")
    };
  }
});

// node_modules/adm-zip/util/index.js
var require_util = __commonJS({
  "node_modules/adm-zip/util/index.js"(exports2, module2) {
    module2.exports = require_utils();
    module2.exports.Constants = require_constants();
    module2.exports.Errors = require_errors();
    module2.exports.FileAttr = require_fattr();
    module2.exports.decoder = require_decoder();
  }
});

// node_modules/adm-zip/headers/entryHeader.js
var require_entryHeader = __commonJS({
  "node_modules/adm-zip/headers/entryHeader.js"(exports2, module2) {
    var Utils = require_util();
    var Constants = Utils.Constants;
    module2.exports = function() {
      var _verMade = 20, _version = 10, _flags = 0, _method = 0, _time = 0, _crc = 0, _compressedSize = 0, _size = 0, _fnameLen = 0, _extraLen = 0, _comLen = 0, _diskStart = 0, _inattr = 0, _attr = 0, _offset = 0;
      _verMade |= Utils.isWin ? 2560 : 768;
      _flags |= Constants.FLG_EFS;
      const _localHeader = {
        extraLen: 0
      };
      const uint32 = (val) => Math.max(0, val) >>> 0;
      const uint16 = (val) => Math.max(0, val) & 65535;
      const uint8 = (val) => Math.max(0, val) & 255;
      _time = Utils.fromDate2DOS(/* @__PURE__ */ new Date());
      return {
        get made() {
          return _verMade;
        },
        set made(val) {
          _verMade = val;
        },
        get version() {
          return _version;
        },
        set version(val) {
          _version = val;
        },
        get flags() {
          return _flags;
        },
        set flags(val) {
          _flags = val;
        },
        get flags_efs() {
          return (_flags & Constants.FLG_EFS) > 0;
        },
        set flags_efs(val) {
          if (val) {
            _flags |= Constants.FLG_EFS;
          } else {
            _flags &= ~Constants.FLG_EFS;
          }
        },
        get flags_desc() {
          return (_flags & Constants.FLG_DESC) > 0;
        },
        set flags_desc(val) {
          if (val) {
            _flags |= Constants.FLG_DESC;
          } else {
            _flags &= ~Constants.FLG_DESC;
          }
        },
        get method() {
          return _method;
        },
        set method(val) {
          switch (val) {
            case Constants.STORED:
              this.version = 10;
            case Constants.DEFLATED:
            default:
              this.version = 20;
          }
          _method = val;
        },
        get time() {
          return Utils.fromDOS2Date(this.timeval);
        },
        set time(val) {
          this.timeval = Utils.fromDate2DOS(val);
        },
        get timeval() {
          return _time;
        },
        set timeval(val) {
          _time = uint32(val);
        },
        get timeHighByte() {
          return uint8(_time >>> 8);
        },
        get crc() {
          return _crc;
        },
        set crc(val) {
          _crc = uint32(val);
        },
        get compressedSize() {
          return _compressedSize;
        },
        set compressedSize(val) {
          _compressedSize = uint32(val);
        },
        get size() {
          return _size;
        },
        set size(val) {
          _size = uint32(val);
        },
        get fileNameLength() {
          return _fnameLen;
        },
        set fileNameLength(val) {
          _fnameLen = val;
        },
        get extraLength() {
          return _extraLen;
        },
        set extraLength(val) {
          _extraLen = val;
        },
        get extraLocalLength() {
          return _localHeader.extraLen;
        },
        set extraLocalLength(val) {
          _localHeader.extraLen = val;
        },
        get commentLength() {
          return _comLen;
        },
        set commentLength(val) {
          _comLen = val;
        },
        get diskNumStart() {
          return _diskStart;
        },
        set diskNumStart(val) {
          _diskStart = uint32(val);
        },
        get inAttr() {
          return _inattr;
        },
        set inAttr(val) {
          _inattr = uint32(val);
        },
        get attr() {
          return _attr;
        },
        set attr(val) {
          _attr = uint32(val);
        },
        // get Unix file permissions
        get fileAttr() {
          return (_attr || 0) >> 16 & 4095;
        },
        get offset() {
          return _offset;
        },
        set offset(val) {
          _offset = uint32(val);
        },
        get encrypted() {
          return (_flags & Constants.FLG_ENC) === Constants.FLG_ENC;
        },
        get centralHeaderSize() {
          return Constants.CENHDR + _fnameLen + _extraLen + _comLen;
        },
        get realDataOffset() {
          return _offset + Constants.LOCHDR + _localHeader.fnameLen + _localHeader.extraLen;
        },
        get localHeader() {
          return _localHeader;
        },
        loadLocalHeaderFromBinary: function(input) {
          var data = input.slice(_offset, _offset + Constants.LOCHDR);
          if (data.readUInt32LE(0) !== Constants.LOCSIG) {
            throw Utils.Errors.INVALID_LOC();
          }
          _localHeader.version = data.readUInt16LE(Constants.LOCVER);
          _localHeader.flags = data.readUInt16LE(Constants.LOCFLG);
          _localHeader.method = data.readUInt16LE(Constants.LOCHOW);
          _localHeader.time = data.readUInt32LE(Constants.LOCTIM);
          _localHeader.crc = data.readUInt32LE(Constants.LOCCRC);
          _localHeader.compressedSize = data.readUInt32LE(Constants.LOCSIZ);
          _localHeader.size = data.readUInt32LE(Constants.LOCLEN);
          _localHeader.fnameLen = data.readUInt16LE(Constants.LOCNAM);
          _localHeader.extraLen = data.readUInt16LE(Constants.LOCEXT);
          const extraStart = _offset + Constants.LOCHDR + _localHeader.fnameLen;
          const extraEnd = extraStart + _localHeader.extraLen;
          return input.slice(extraStart, extraEnd);
        },
        loadFromBinary: function(data) {
          if (data.length !== Constants.CENHDR || data.readUInt32LE(0) !== Constants.CENSIG) {
            throw Utils.Errors.INVALID_CEN();
          }
          _verMade = data.readUInt16LE(Constants.CENVEM);
          _version = data.readUInt16LE(Constants.CENVER);
          _flags = data.readUInt16LE(Constants.CENFLG);
          _method = data.readUInt16LE(Constants.CENHOW);
          _time = data.readUInt32LE(Constants.CENTIM);
          _crc = data.readUInt32LE(Constants.CENCRC);
          _compressedSize = data.readUInt32LE(Constants.CENSIZ);
          _size = data.readUInt32LE(Constants.CENLEN);
          _fnameLen = data.readUInt16LE(Constants.CENNAM);
          _extraLen = data.readUInt16LE(Constants.CENEXT);
          _comLen = data.readUInt16LE(Constants.CENCOM);
          _diskStart = data.readUInt16LE(Constants.CENDSK);
          _inattr = data.readUInt16LE(Constants.CENATT);
          _attr = data.readUInt32LE(Constants.CENATX);
          _offset = data.readUInt32LE(Constants.CENOFF);
        },
        localHeaderToBinary: function() {
          var data = Buffer.alloc(Constants.LOCHDR);
          data.writeUInt32LE(Constants.LOCSIG, 0);
          data.writeUInt16LE(_version, Constants.LOCVER);
          data.writeUInt16LE(_flags, Constants.LOCFLG);
          data.writeUInt16LE(_method, Constants.LOCHOW);
          data.writeUInt32LE(_time, Constants.LOCTIM);
          data.writeUInt32LE(_crc, Constants.LOCCRC);
          data.writeUInt32LE(_compressedSize, Constants.LOCSIZ);
          data.writeUInt32LE(_size, Constants.LOCLEN);
          data.writeUInt16LE(_fnameLen, Constants.LOCNAM);
          data.writeUInt16LE(_localHeader.extraLen, Constants.LOCEXT);
          return data;
        },
        centralHeaderToBinary: function() {
          var data = Buffer.alloc(Constants.CENHDR + _fnameLen + _extraLen + _comLen);
          data.writeUInt32LE(Constants.CENSIG, 0);
          data.writeUInt16LE(_verMade, Constants.CENVEM);
          data.writeUInt16LE(_version, Constants.CENVER);
          data.writeUInt16LE(_flags, Constants.CENFLG);
          data.writeUInt16LE(_method, Constants.CENHOW);
          data.writeUInt32LE(_time, Constants.CENTIM);
          data.writeUInt32LE(_crc, Constants.CENCRC);
          data.writeUInt32LE(_compressedSize, Constants.CENSIZ);
          data.writeUInt32LE(_size, Constants.CENLEN);
          data.writeUInt16LE(_fnameLen, Constants.CENNAM);
          data.writeUInt16LE(_extraLen, Constants.CENEXT);
          data.writeUInt16LE(_comLen, Constants.CENCOM);
          data.writeUInt16LE(_diskStart, Constants.CENDSK);
          data.writeUInt16LE(_inattr, Constants.CENATT);
          data.writeUInt32LE(_attr, Constants.CENATX);
          data.writeUInt32LE(_offset, Constants.CENOFF);
          return data;
        },
        toJSON: function() {
          const bytes = function(nr) {
            return nr + " bytes";
          };
          return {
            made: _verMade,
            version: _version,
            flags: _flags,
            method: Utils.methodToString(_method),
            time: this.time,
            crc: "0x" + _crc.toString(16).toUpperCase(),
            compressedSize: bytes(_compressedSize),
            size: bytes(_size),
            fileNameLength: bytes(_fnameLen),
            extraLength: bytes(_extraLen),
            commentLength: bytes(_comLen),
            diskNumStart: _diskStart,
            inAttr: _inattr,
            attr: _attr,
            offset: _offset,
            centralHeaderSize: bytes(Constants.CENHDR + _fnameLen + _extraLen + _comLen)
          };
        },
        toString: function() {
          return JSON.stringify(this.toJSON(), null, "	");
        }
      };
    };
  }
});

// node_modules/adm-zip/headers/mainHeader.js
var require_mainHeader = __commonJS({
  "node_modules/adm-zip/headers/mainHeader.js"(exports2, module2) {
    var Utils = require_util();
    var Constants = Utils.Constants;
    module2.exports = function() {
      var _volumeEntries = 0, _totalEntries = 0, _size = 0, _offset = 0, _commentLength = 0;
      return {
        get diskEntries() {
          return _volumeEntries;
        },
        set diskEntries(val) {
          _volumeEntries = _totalEntries = val;
        },
        get totalEntries() {
          return _totalEntries;
        },
        set totalEntries(val) {
          _totalEntries = _volumeEntries = val;
        },
        get size() {
          return _size;
        },
        set size(val) {
          _size = val;
        },
        get offset() {
          return _offset;
        },
        set offset(val) {
          _offset = val;
        },
        get commentLength() {
          return _commentLength;
        },
        set commentLength(val) {
          _commentLength = val;
        },
        get mainHeaderSize() {
          return Constants.ENDHDR + _commentLength;
        },
        loadFromBinary: function(data) {
          if ((data.length !== Constants.ENDHDR || data.readUInt32LE(0) !== Constants.ENDSIG) && (data.length < Constants.ZIP64HDR || data.readUInt32LE(0) !== Constants.ZIP64SIG)) {
            throw Utils.Errors.INVALID_END();
          }
          if (data.readUInt32LE(0) === Constants.ENDSIG) {
            _volumeEntries = data.readUInt16LE(Constants.ENDSUB);
            _totalEntries = data.readUInt16LE(Constants.ENDTOT);
            _size = data.readUInt32LE(Constants.ENDSIZ);
            _offset = data.readUInt32LE(Constants.ENDOFF);
            _commentLength = data.readUInt16LE(Constants.ENDCOM);
          } else {
            _volumeEntries = Utils.readBigUInt64LE(data, Constants.ZIP64SUB);
            _totalEntries = Utils.readBigUInt64LE(data, Constants.ZIP64TOT);
            _size = Utils.readBigUInt64LE(data, Constants.ZIP64SIZE);
            _offset = Utils.readBigUInt64LE(data, Constants.ZIP64OFF);
            _commentLength = 0;
          }
        },
        toBinary: function() {
          var b = Buffer.alloc(Constants.ENDHDR + _commentLength);
          b.writeUInt32LE(Constants.ENDSIG, 0);
          b.writeUInt32LE(0, 4);
          b.writeUInt16LE(_volumeEntries, Constants.ENDSUB);
          b.writeUInt16LE(_totalEntries, Constants.ENDTOT);
          b.writeUInt32LE(_size, Constants.ENDSIZ);
          b.writeUInt32LE(_offset, Constants.ENDOFF);
          b.writeUInt16LE(_commentLength, Constants.ENDCOM);
          b.fill(" ", Constants.ENDHDR);
          return b;
        },
        toJSON: function() {
          const offset = function(nr, len) {
            let offs = nr.toString(16).toUpperCase();
            while (offs.length < len) offs = "0" + offs;
            return "0x" + offs;
          };
          return {
            diskEntries: _volumeEntries,
            totalEntries: _totalEntries,
            size: _size + " bytes",
            offset: offset(_offset, 4),
            commentLength: _commentLength
          };
        },
        toString: function() {
          return JSON.stringify(this.toJSON(), null, "	");
        }
      };
    };
  }
});

// node_modules/adm-zip/headers/index.js
var require_headers = __commonJS({
  "node_modules/adm-zip/headers/index.js"(exports2) {
    exports2.EntryHeader = require_entryHeader();
    exports2.MainHeader = require_mainHeader();
  }
});

// node_modules/adm-zip/methods/deflater.js
var require_deflater = __commonJS({
  "node_modules/adm-zip/methods/deflater.js"(exports2, module2) {
    module2.exports = function(inbuf) {
      var zlib = require("zlib");
      var opts = { chunkSize: (parseInt(inbuf.length / 1024) + 1) * 1024 };
      return {
        deflate: function() {
          return zlib.deflateRawSync(inbuf, opts);
        },
        deflateAsync: function(callback) {
          var tmp = zlib.createDeflateRaw(opts), parts = [], total = 0;
          tmp.on("data", function(data) {
            parts.push(data);
            total += data.length;
          });
          tmp.on("end", function() {
            var buf = Buffer.alloc(total), written = 0;
            buf.fill(0);
            for (var i = 0; i < parts.length; i++) {
              var part = parts[i];
              part.copy(buf, written);
              written += part.length;
            }
            callback && callback(buf);
          });
          tmp.end(inbuf);
        }
      };
    };
  }
});

// node_modules/adm-zip/methods/inflater.js
var require_inflater = __commonJS({
  "node_modules/adm-zip/methods/inflater.js"(exports2, module2) {
    var version = +(process.versions ? process.versions.node : "").split(".")[0] || 0;
    module2.exports = function(inbuf, expectedLength) {
      var zlib = require("zlib");
      const option = version >= 15 && expectedLength > 0 ? { maxOutputLength: expectedLength } : {};
      return {
        inflate: function() {
          return zlib.inflateRawSync(inbuf, option);
        },
        inflateAsync: function(callback) {
          var tmp = zlib.createInflateRaw(option), parts = [], total = 0;
          tmp.on("data", function(data) {
            parts.push(data);
            total += data.length;
          });
          tmp.on("end", function() {
            var buf = Buffer.alloc(total), written = 0;
            buf.fill(0);
            for (var i = 0; i < parts.length; i++) {
              var part = parts[i];
              part.copy(buf, written);
              written += part.length;
            }
            callback && callback(buf);
          });
          tmp.end(inbuf);
        }
      };
    };
  }
});

// node_modules/adm-zip/methods/zipcrypto.js
var require_zipcrypto = __commonJS({
  "node_modules/adm-zip/methods/zipcrypto.js"(exports2, module2) {
    "use strict";
    var { randomFillSync } = require("crypto");
    var Errors = require_errors();
    var crctable = new Uint32Array(256).map((t, crc) => {
      for (let j = 0; j < 8; j++) {
        if (0 !== (crc & 1)) {
          crc = crc >>> 1 ^ 3988292384;
        } else {
          crc >>>= 1;
        }
      }
      return crc >>> 0;
    });
    var uMul = (a, b) => Math.imul(a, b) >>> 0;
    var crc32update = (pCrc32, bval) => {
      return crctable[(pCrc32 ^ bval) & 255] ^ pCrc32 >>> 8;
    };
    var genSalt = () => {
      if ("function" === typeof randomFillSync) {
        return randomFillSync(Buffer.alloc(12));
      } else {
        return genSalt.node();
      }
    };
    genSalt.node = () => {
      const salt = Buffer.alloc(12);
      const len = salt.length;
      for (let i = 0; i < len; i++) salt[i] = Math.random() * 256 & 255;
      return salt;
    };
    var config = {
      genSalt
    };
    function Initkeys(pw) {
      const pass = Buffer.isBuffer(pw) ? pw : Buffer.from(pw);
      this.keys = new Uint32Array([305419896, 591751049, 878082192]);
      for (let i = 0; i < pass.length; i++) {
        this.updateKeys(pass[i]);
      }
    }
    Initkeys.prototype.updateKeys = function(byteValue) {
      const keys = this.keys;
      keys[0] = crc32update(keys[0], byteValue);
      keys[1] += keys[0] & 255;
      keys[1] = uMul(keys[1], 134775813) + 1;
      keys[2] = crc32update(keys[2], keys[1] >>> 24);
      return byteValue;
    };
    Initkeys.prototype.next = function() {
      const k = (this.keys[2] | 2) >>> 0;
      return uMul(k, k ^ 1) >> 8 & 255;
    };
    function make_decrypter(pwd) {
      const keys = new Initkeys(pwd);
      return function(data) {
        const result = Buffer.alloc(data.length);
        let pos = 0;
        for (let c of data) {
          result[pos++] = keys.updateKeys(c ^ keys.next());
        }
        return result;
      };
    }
    function make_encrypter(pwd) {
      const keys = new Initkeys(pwd);
      return function(data, result, pos = 0) {
        if (!result) result = Buffer.alloc(data.length);
        for (let c of data) {
          const k = keys.next();
          result[pos++] = c ^ k;
          keys.updateKeys(c);
        }
        return result;
      };
    }
    function decrypt(data, header, pwd) {
      if (!data || !Buffer.isBuffer(data) || data.length < 12) {
        return Buffer.alloc(0);
      }
      const decrypter = make_decrypter(pwd);
      const salt = decrypter(data.slice(0, 12));
      const verifyByte = (header.flags & 8) === 8 ? header.timeHighByte : header.crc >>> 24;
      if (salt[11] !== verifyByte) {
        throw Errors.WRONG_PASSWORD();
      }
      return decrypter(data.slice(12));
    }
    function _salter(data) {
      if (Buffer.isBuffer(data) && data.length >= 12) {
        config.genSalt = function() {
          return data.slice(0, 12);
        };
      } else if (data === "node") {
        config.genSalt = genSalt.node;
      } else {
        config.genSalt = genSalt;
      }
    }
    function encrypt(data, header, pwd, oldlike = false) {
      if (data == null) data = Buffer.alloc(0);
      if (!Buffer.isBuffer(data)) data = Buffer.from(data.toString());
      const encrypter = make_encrypter(pwd);
      const salt = config.genSalt();
      salt[11] = header.crc >>> 24 & 255;
      if (oldlike) salt[10] = header.crc >>> 16 & 255;
      const result = Buffer.alloc(data.length + 12);
      encrypter(salt, result);
      return encrypter(data, result, 12);
    }
    module2.exports = { decrypt, encrypt, _salter };
  }
});

// node_modules/adm-zip/methods/index.js
var require_methods = __commonJS({
  "node_modules/adm-zip/methods/index.js"(exports2) {
    exports2.Deflater = require_deflater();
    exports2.Inflater = require_inflater();
    exports2.ZipCrypto = require_zipcrypto();
  }
});

// node_modules/adm-zip/zipEntry.js
var require_zipEntry = __commonJS({
  "node_modules/adm-zip/zipEntry.js"(exports2, module2) {
    var Utils = require_util();
    var Headers = require_headers();
    var Constants = Utils.Constants;
    var Methods = require_methods();
    module2.exports = function(options, input) {
      var _centralHeader = new Headers.EntryHeader(), _entryName = Buffer.alloc(0), _comment = Buffer.alloc(0), _isDirectory = false, uncompressedData = null, _extra = Buffer.alloc(0), _extralocal = Buffer.alloc(0), _efs = true;
      const opts = options;
      const decoder = typeof opts.decoder === "object" ? opts.decoder : Utils.decoder;
      _efs = decoder.hasOwnProperty("efs") ? decoder.efs : false;
      function getCompressedDataFromZip() {
        if (!input || !(input instanceof Uint8Array)) {
          return Buffer.alloc(0);
        }
        _extralocal = _centralHeader.loadLocalHeaderFromBinary(input);
        return input.slice(_centralHeader.realDataOffset, _centralHeader.realDataOffset + _centralHeader.compressedSize);
      }
      function crc32OK(data) {
        if (!_centralHeader.flags_desc) {
          if (Utils.crc32(data) !== _centralHeader.localHeader.crc) {
            return false;
          }
        } else {
          const descriptor = {};
          const dataEndOffset = _centralHeader.realDataOffset + _centralHeader.compressedSize;
          if (input.readUInt32LE(dataEndOffset) == Constants.LOCSIG || input.readUInt32LE(dataEndOffset) == Constants.CENSIG) {
            throw Utils.Errors.DESCRIPTOR_NOT_EXIST();
          }
          if (input.readUInt32LE(dataEndOffset) == Constants.EXTSIG) {
            descriptor.crc = input.readUInt32LE(dataEndOffset + Constants.EXTCRC);
            descriptor.compressedSize = input.readUInt32LE(dataEndOffset + Constants.EXTSIZ);
            descriptor.size = input.readUInt32LE(dataEndOffset + Constants.EXTLEN);
          } else if (input.readUInt16LE(dataEndOffset + 12) === 19280) {
            descriptor.crc = input.readUInt32LE(dataEndOffset + Constants.EXTCRC - 4);
            descriptor.compressedSize = input.readUInt32LE(dataEndOffset + Constants.EXTSIZ - 4);
            descriptor.size = input.readUInt32LE(dataEndOffset + Constants.EXTLEN - 4);
          } else {
            throw Utils.Errors.DESCRIPTOR_UNKNOWN();
          }
          if (descriptor.compressedSize !== _centralHeader.compressedSize || descriptor.size !== _centralHeader.size || descriptor.crc !== _centralHeader.crc) {
            throw Utils.Errors.DESCRIPTOR_FAULTY();
          }
          if (Utils.crc32(data) !== descriptor.crc) {
            return false;
          }
        }
        return true;
      }
      function decompress(async, callback, pass) {
        if (typeof callback === "undefined" && typeof async === "string") {
          pass = async;
          async = void 0;
        }
        if (_isDirectory) {
          if (async && callback) {
            callback(Buffer.alloc(0), Utils.Errors.DIRECTORY_CONTENT_ERROR());
          }
          return Buffer.alloc(0);
        }
        var compressedData = getCompressedDataFromZip();
        if (compressedData.length === 0) {
          if (async && callback) callback(compressedData);
          return compressedData;
        }
        if (_centralHeader.encrypted) {
          if ("string" !== typeof pass && !Buffer.isBuffer(pass)) {
            throw Utils.Errors.INVALID_PASS_PARAM();
          }
          compressedData = Methods.ZipCrypto.decrypt(compressedData, _centralHeader, pass);
        }
        var data = Buffer.alloc(_centralHeader.size);
        switch (_centralHeader.method) {
          case Utils.Constants.STORED:
            compressedData.copy(data);
            if (!crc32OK(data)) {
              if (async && callback) callback(data, Utils.Errors.BAD_CRC());
              throw Utils.Errors.BAD_CRC();
            } else {
              if (async && callback) callback(data);
              return data;
            }
          case Utils.Constants.DEFLATED:
            var inflater = new Methods.Inflater(compressedData, _centralHeader.size);
            if (!async) {
              const result = inflater.inflate(data);
              result.copy(data, 0);
              if (!crc32OK(data)) {
                throw Utils.Errors.BAD_CRC(`"${decoder.decode(_entryName)}"`);
              }
              return data;
            } else {
              inflater.inflateAsync(function(result) {
                result.copy(result, 0);
                if (callback) {
                  if (!crc32OK(result)) {
                    callback(result, Utils.Errors.BAD_CRC());
                  } else {
                    callback(result);
                  }
                }
              });
            }
            break;
          default:
            if (async && callback) callback(Buffer.alloc(0), Utils.Errors.UNKNOWN_METHOD());
            throw Utils.Errors.UNKNOWN_METHOD();
        }
      }
      function compress(async, callback) {
        if ((!uncompressedData || !uncompressedData.length) && Buffer.isBuffer(input)) {
          if (async && callback) callback(getCompressedDataFromZip());
          return getCompressedDataFromZip();
        }
        if (uncompressedData.length && !_isDirectory) {
          var compressedData;
          switch (_centralHeader.method) {
            case Utils.Constants.STORED:
              _centralHeader.compressedSize = _centralHeader.size;
              compressedData = Buffer.alloc(uncompressedData.length);
              uncompressedData.copy(compressedData);
              if (async && callback) callback(compressedData);
              return compressedData;
            default:
            case Utils.Constants.DEFLATED:
              var deflater = new Methods.Deflater(uncompressedData);
              if (!async) {
                var deflated = deflater.deflate();
                _centralHeader.compressedSize = deflated.length;
                return deflated;
              } else {
                deflater.deflateAsync(function(data) {
                  compressedData = Buffer.alloc(data.length);
                  _centralHeader.compressedSize = data.length;
                  data.copy(compressedData);
                  callback && callback(compressedData);
                });
              }
              deflater = null;
              break;
          }
        } else if (async && callback) {
          callback(Buffer.alloc(0));
        } else {
          return Buffer.alloc(0);
        }
      }
      function readUInt64LE(buffer, offset) {
        return (buffer.readUInt32LE(offset + 4) << 4) + buffer.readUInt32LE(offset);
      }
      function parseExtra(data) {
        try {
          var offset = 0;
          var signature, size, part;
          while (offset + 4 < data.length) {
            signature = data.readUInt16LE(offset);
            offset += 2;
            size = data.readUInt16LE(offset);
            offset += 2;
            part = data.slice(offset, offset + size);
            offset += size;
            if (Constants.ID_ZIP64 === signature) {
              parseZip64ExtendedInformation(part);
            }
          }
        } catch (error) {
          throw Utils.Errors.EXTRA_FIELD_PARSE_ERROR();
        }
      }
      function parseZip64ExtendedInformation(data) {
        var size, compressedSize, offset, diskNumStart;
        if (data.length >= Constants.EF_ZIP64_SCOMP) {
          size = readUInt64LE(data, Constants.EF_ZIP64_SUNCOMP);
          if (_centralHeader.size === Constants.EF_ZIP64_OR_32) {
            _centralHeader.size = size;
          }
        }
        if (data.length >= Constants.EF_ZIP64_RHO) {
          compressedSize = readUInt64LE(data, Constants.EF_ZIP64_SCOMP);
          if (_centralHeader.compressedSize === Constants.EF_ZIP64_OR_32) {
            _centralHeader.compressedSize = compressedSize;
          }
        }
        if (data.length >= Constants.EF_ZIP64_DSN) {
          offset = readUInt64LE(data, Constants.EF_ZIP64_RHO);
          if (_centralHeader.offset === Constants.EF_ZIP64_OR_32) {
            _centralHeader.offset = offset;
          }
        }
        if (data.length >= Constants.EF_ZIP64_DSN + 4) {
          diskNumStart = data.readUInt32LE(Constants.EF_ZIP64_DSN);
          if (_centralHeader.diskNumStart === Constants.EF_ZIP64_OR_16) {
            _centralHeader.diskNumStart = diskNumStart;
          }
        }
      }
      return {
        get entryName() {
          return decoder.decode(_entryName);
        },
        get rawEntryName() {
          return _entryName;
        },
        set entryName(val) {
          _entryName = Utils.toBuffer(val, decoder.encode);
          var lastChar = _entryName[_entryName.length - 1];
          _isDirectory = lastChar === 47 || lastChar === 92;
          _centralHeader.fileNameLength = _entryName.length;
        },
        get efs() {
          if (typeof _efs === "function") {
            return _efs(this.entryName);
          } else {
            return _efs;
          }
        },
        get extra() {
          return _extra;
        },
        set extra(val) {
          _extra = val;
          _centralHeader.extraLength = val.length;
          parseExtra(val);
        },
        get comment() {
          return decoder.decode(_comment);
        },
        set comment(val) {
          _comment = Utils.toBuffer(val, decoder.encode);
          _centralHeader.commentLength = _comment.length;
          if (_comment.length > 65535) throw Utils.Errors.COMMENT_TOO_LONG();
        },
        get name() {
          var n = decoder.decode(_entryName);
          return _isDirectory ? n.substr(n.length - 1).split("/").pop() : n.split("/").pop();
        },
        get isDirectory() {
          return _isDirectory;
        },
        getCompressedData: function() {
          return compress(false, null);
        },
        getCompressedDataAsync: function(callback) {
          compress(true, callback);
        },
        setData: function(value) {
          uncompressedData = Utils.toBuffer(value, Utils.decoder.encode);
          if (!_isDirectory && uncompressedData.length) {
            _centralHeader.size = uncompressedData.length;
            _centralHeader.method = Utils.Constants.DEFLATED;
            _centralHeader.crc = Utils.crc32(value);
            _centralHeader.changed = true;
          } else {
            _centralHeader.method = Utils.Constants.STORED;
          }
        },
        getData: function(pass) {
          if (_centralHeader.changed) {
            return uncompressedData;
          } else {
            return decompress(false, null, pass);
          }
        },
        getDataAsync: function(callback, pass) {
          if (_centralHeader.changed) {
            callback(uncompressedData);
          } else {
            decompress(true, callback, pass);
          }
        },
        set attr(attr) {
          _centralHeader.attr = attr;
        },
        get attr() {
          return _centralHeader.attr;
        },
        set header(data) {
          _centralHeader.loadFromBinary(data);
        },
        get header() {
          return _centralHeader;
        },
        packCentralHeader: function() {
          _centralHeader.flags_efs = this.efs;
          _centralHeader.extraLength = _extra.length;
          var header = _centralHeader.centralHeaderToBinary();
          var addpos = Utils.Constants.CENHDR;
          _entryName.copy(header, addpos);
          addpos += _entryName.length;
          _extra.copy(header, addpos);
          addpos += _centralHeader.extraLength;
          _comment.copy(header, addpos);
          return header;
        },
        packLocalHeader: function() {
          let addpos = 0;
          _centralHeader.flags_efs = this.efs;
          _centralHeader.extraLocalLength = _extralocal.length;
          const localHeaderBuf = _centralHeader.localHeaderToBinary();
          const localHeader = Buffer.alloc(localHeaderBuf.length + _entryName.length + _centralHeader.extraLocalLength);
          localHeaderBuf.copy(localHeader, addpos);
          addpos += localHeaderBuf.length;
          _entryName.copy(localHeader, addpos);
          addpos += _entryName.length;
          _extralocal.copy(localHeader, addpos);
          addpos += _extralocal.length;
          return localHeader;
        },
        toJSON: function() {
          const bytes = function(nr) {
            return "<" + (nr && nr.length + " bytes buffer" || "null") + ">";
          };
          return {
            entryName: this.entryName,
            name: this.name,
            comment: this.comment,
            isDirectory: this.isDirectory,
            header: _centralHeader.toJSON(),
            compressedData: bytes(input),
            data: bytes(uncompressedData)
          };
        },
        toString: function() {
          return JSON.stringify(this.toJSON(), null, "	");
        }
      };
    };
  }
});

// node_modules/adm-zip/zipFile.js
var require_zipFile = __commonJS({
  "node_modules/adm-zip/zipFile.js"(exports2, module2) {
    var ZipEntry = require_zipEntry();
    var Headers = require_headers();
    var Utils = require_util();
    module2.exports = function(inBuffer, options) {
      var entryList = [], entryTable = {}, _comment = Buffer.alloc(0), mainHeader = new Headers.MainHeader(), loadedEntries = false;
      var password = null;
      const temporary = /* @__PURE__ */ new Set();
      const opts = options;
      const { noSort, decoder } = opts;
      if (inBuffer) {
        readMainHeader(opts.readEntries);
      } else {
        loadedEntries = true;
      }
      function makeTemporaryFolders() {
        const foldersList = /* @__PURE__ */ new Set();
        for (const elem of Object.keys(entryTable)) {
          const elements = elem.split("/");
          elements.pop();
          if (!elements.length) continue;
          for (let i = 0; i < elements.length; i++) {
            const sub = elements.slice(0, i + 1).join("/") + "/";
            foldersList.add(sub);
          }
        }
        for (const elem of foldersList) {
          if (!(elem in entryTable)) {
            const tempfolder = new ZipEntry(opts);
            tempfolder.entryName = elem;
            tempfolder.attr = 16;
            tempfolder.temporary = true;
            entryList.push(tempfolder);
            entryTable[tempfolder.entryName] = tempfolder;
            temporary.add(tempfolder);
          }
        }
      }
      function readEntries() {
        loadedEntries = true;
        entryTable = {};
        if (mainHeader.diskEntries > (inBuffer.length - mainHeader.offset) / Utils.Constants.CENHDR) {
          throw Utils.Errors.DISK_ENTRY_TOO_LARGE();
        }
        entryList = new Array(mainHeader.diskEntries);
        var index = mainHeader.offset;
        for (var i = 0; i < entryList.length; i++) {
          var tmp = index, entry = new ZipEntry(opts, inBuffer);
          entry.header = inBuffer.slice(tmp, tmp += Utils.Constants.CENHDR);
          entry.entryName = inBuffer.slice(tmp, tmp += entry.header.fileNameLength);
          if (entry.header.extraLength) {
            entry.extra = inBuffer.slice(tmp, tmp += entry.header.extraLength);
          }
          if (entry.header.commentLength) entry.comment = inBuffer.slice(tmp, tmp + entry.header.commentLength);
          index += entry.header.centralHeaderSize;
          entryList[i] = entry;
          entryTable[entry.entryName] = entry;
        }
        temporary.clear();
        makeTemporaryFolders();
      }
      function readMainHeader(readNow) {
        var i = inBuffer.length - Utils.Constants.ENDHDR, max = Math.max(0, i - 65535), n = max, endStart = inBuffer.length, endOffset = -1, commentEnd = 0;
        const trailingSpace = typeof opts.trailingSpace === "boolean" ? opts.trailingSpace : false;
        if (trailingSpace) max = 0;
        for (i; i >= n; i--) {
          if (inBuffer[i] !== 80) continue;
          if (inBuffer.readUInt32LE(i) === Utils.Constants.ENDSIG) {
            endOffset = i;
            commentEnd = i;
            endStart = i + Utils.Constants.ENDHDR;
            n = i - Utils.Constants.END64HDR;
            continue;
          }
          if (inBuffer.readUInt32LE(i) === Utils.Constants.END64SIG) {
            n = max;
            continue;
          }
          if (inBuffer.readUInt32LE(i) === Utils.Constants.ZIP64SIG) {
            endOffset = i;
            endStart = i + Utils.readBigUInt64LE(inBuffer, i + Utils.Constants.ZIP64SIZE) + Utils.Constants.ZIP64LEAD;
            break;
          }
        }
        if (endOffset == -1) throw Utils.Errors.INVALID_FORMAT();
        mainHeader.loadFromBinary(inBuffer.slice(endOffset, endStart));
        if (mainHeader.commentLength) {
          _comment = inBuffer.slice(commentEnd + Utils.Constants.ENDHDR);
        }
        if (readNow) readEntries();
      }
      function sortEntries() {
        if (entryList.length > 1 && !noSort) {
          entryList.sort((a, b) => a.entryName.toLowerCase().localeCompare(b.entryName.toLowerCase()));
        }
      }
      return {
        /**
         * Returns an array of ZipEntry objects existent in the current opened archive
         * @return Array
         */
        get entries() {
          if (!loadedEntries) {
            readEntries();
          }
          return entryList.filter((e) => !temporary.has(e));
        },
        /**
         * Archive comment
         * @return {String}
         */
        get comment() {
          return decoder.decode(_comment);
        },
        set comment(val) {
          _comment = Utils.toBuffer(val, decoder.encode);
          mainHeader.commentLength = _comment.length;
        },
        getEntryCount: function() {
          if (!loadedEntries) {
            return mainHeader.diskEntries;
          }
          return entryList.length;
        },
        forEach: function(callback) {
          this.entries.forEach(callback);
        },
        /**
         * Returns a reference to the entry with the given name or null if entry is inexistent
         *
         * @param entryName
         * @return ZipEntry
         */
        getEntry: function(entryName) {
          if (!loadedEntries) {
            readEntries();
          }
          return entryTable[entryName] || null;
        },
        /**
         * Adds the given entry to the entry list
         *
         * @param entry
         */
        setEntry: function(entry) {
          if (!loadedEntries) {
            readEntries();
          }
          entryList.push(entry);
          entryTable[entry.entryName] = entry;
          mainHeader.totalEntries = entryList.length;
        },
        /**
         * Removes the file with the given name from the entry list.
         *
         * If the entry is a directory, then all nested files and directories will be removed
         * @param entryName
         * @returns {void}
         */
        deleteFile: function(entryName, withsubfolders = true) {
          if (!loadedEntries) {
            readEntries();
          }
          const entry = entryTable[entryName];
          const list = this.getEntryChildren(entry, withsubfolders).map((child) => child.entryName);
          list.forEach(this.deleteEntry);
        },
        /**
         * Removes the entry with the given name from the entry list.
         *
         * @param {string} entryName
         * @returns {void}
         */
        deleteEntry: function(entryName) {
          if (!loadedEntries) {
            readEntries();
          }
          const entry = entryTable[entryName];
          const index = entryList.indexOf(entry);
          if (index >= 0) {
            entryList.splice(index, 1);
            delete entryTable[entryName];
            mainHeader.totalEntries = entryList.length;
          }
        },
        /**
         *  Iterates and returns all nested files and directories of the given entry
         *
         * @param entry
         * @return Array
         */
        getEntryChildren: function(entry, subfolders = true) {
          if (!loadedEntries) {
            readEntries();
          }
          if (typeof entry === "object") {
            if (entry.isDirectory && subfolders) {
              const list = [];
              const name = entry.entryName;
              for (const zipEntry of entryList) {
                if (zipEntry.entryName.startsWith(name)) {
                  list.push(zipEntry);
                }
              }
              return list;
            } else {
              return [entry];
            }
          }
          return [];
        },
        /**
         *  How many child elements entry has
         *
         * @param {ZipEntry} entry
         * @return {integer}
         */
        getChildCount: function(entry) {
          if (entry && entry.isDirectory) {
            const list = this.getEntryChildren(entry);
            return list.includes(entry) ? list.length - 1 : list.length;
          }
          return 0;
        },
        /**
         * Returns the zip file
         *
         * @return Buffer
         */
        compressToBuffer: function() {
          if (!loadedEntries) {
            readEntries();
          }
          sortEntries();
          const dataBlock = [];
          const headerBlocks = [];
          let totalSize = 0;
          let dindex = 0;
          mainHeader.size = 0;
          mainHeader.offset = 0;
          let totalEntries = 0;
          for (const entry of this.entries) {
            const compressedData = entry.getCompressedData();
            entry.header.offset = dindex;
            const localHeader = entry.packLocalHeader();
            const dataLength = localHeader.length + compressedData.length;
            dindex += dataLength;
            dataBlock.push(localHeader);
            dataBlock.push(compressedData);
            const centralHeader = entry.packCentralHeader();
            headerBlocks.push(centralHeader);
            mainHeader.size += centralHeader.length;
            totalSize += dataLength + centralHeader.length;
            totalEntries++;
          }
          totalSize += mainHeader.mainHeaderSize;
          mainHeader.offset = dindex;
          mainHeader.totalEntries = totalEntries;
          dindex = 0;
          const outBuffer = Buffer.alloc(totalSize);
          for (const content of dataBlock) {
            content.copy(outBuffer, dindex);
            dindex += content.length;
          }
          for (const content of headerBlocks) {
            content.copy(outBuffer, dindex);
            dindex += content.length;
          }
          const mh = mainHeader.toBinary();
          if (_comment) {
            _comment.copy(mh, Utils.Constants.ENDHDR);
          }
          mh.copy(outBuffer, dindex);
          inBuffer = outBuffer;
          loadedEntries = false;
          return outBuffer;
        },
        toAsyncBuffer: function(onSuccess, onFail, onItemStart, onItemEnd) {
          try {
            if (!loadedEntries) {
              readEntries();
            }
            sortEntries();
            const dataBlock = [];
            const centralHeaders = [];
            let totalSize = 0;
            let dindex = 0;
            let totalEntries = 0;
            mainHeader.size = 0;
            mainHeader.offset = 0;
            const compress2Buffer = function(entryLists) {
              if (entryLists.length > 0) {
                const entry = entryLists.shift();
                const name = entry.entryName + entry.extra.toString();
                if (onItemStart) onItemStart(name);
                entry.getCompressedDataAsync(function(compressedData) {
                  if (onItemEnd) onItemEnd(name);
                  entry.header.offset = dindex;
                  const localHeader = entry.packLocalHeader();
                  const dataLength = localHeader.length + compressedData.length;
                  dindex += dataLength;
                  dataBlock.push(localHeader);
                  dataBlock.push(compressedData);
                  const centalHeader = entry.packCentralHeader();
                  centralHeaders.push(centalHeader);
                  mainHeader.size += centalHeader.length;
                  totalSize += dataLength + centalHeader.length;
                  totalEntries++;
                  compress2Buffer(entryLists);
                });
              } else {
                totalSize += mainHeader.mainHeaderSize;
                mainHeader.offset = dindex;
                mainHeader.totalEntries = totalEntries;
                dindex = 0;
                const outBuffer = Buffer.alloc(totalSize);
                dataBlock.forEach(function(content) {
                  content.copy(outBuffer, dindex);
                  dindex += content.length;
                });
                centralHeaders.forEach(function(content) {
                  content.copy(outBuffer, dindex);
                  dindex += content.length;
                });
                const mh = mainHeader.toBinary();
                if (_comment) {
                  _comment.copy(mh, Utils.Constants.ENDHDR);
                }
                mh.copy(outBuffer, dindex);
                inBuffer = outBuffer;
                loadedEntries = false;
                onSuccess(outBuffer);
              }
            };
            compress2Buffer(Array.from(this.entries));
          } catch (e) {
            onFail(e);
          }
        }
      };
    };
  }
});

// node_modules/adm-zip/adm-zip.js
var require_adm_zip = __commonJS({
  "node_modules/adm-zip/adm-zip.js"(exports2, module2) {
    var Utils = require_util();
    var pth = require("path");
    var ZipEntry = require_zipEntry();
    var ZipFile = require_zipFile();
    var get_Bool = (...val) => Utils.findLast(val, (c) => typeof c === "boolean");
    var get_Str = (...val) => Utils.findLast(val, (c) => typeof c === "string");
    var get_Fun = (...val) => Utils.findLast(val, (c) => typeof c === "function");
    var defaultOptions = {
      // option "noSort" : if true it disables files sorting
      noSort: false,
      // read entries during load (initial loading may be slower)
      readEntries: false,
      // default method is none
      method: Utils.Constants.NONE,
      // file system
      fs: null
    };
    module2.exports = function(input, options) {
      let inBuffer = null;
      const opts = Object.assign(/* @__PURE__ */ Object.create(null), defaultOptions);
      if (input && "object" === typeof input) {
        if (!(input instanceof Uint8Array)) {
          Object.assign(opts, input);
          input = opts.input ? opts.input : void 0;
          if (opts.input) delete opts.input;
        }
        if (Buffer.isBuffer(input)) {
          inBuffer = input;
          opts.method = Utils.Constants.BUFFER;
          input = void 0;
        }
      }
      Object.assign(opts, options);
      const filetools = new Utils(opts);
      if (typeof opts.decoder !== "object" || typeof opts.decoder.encode !== "function" || typeof opts.decoder.decode !== "function") {
        opts.decoder = Utils.decoder;
      }
      if (input && "string" === typeof input) {
        if (filetools.fs.existsSync(input)) {
          opts.method = Utils.Constants.FILE;
          opts.filename = input;
          inBuffer = filetools.fs.readFileSync(input);
        } else {
          throw Utils.Errors.INVALID_FILENAME();
        }
      }
      const _zip = new ZipFile(inBuffer, opts);
      const { canonical, sanitize, zipnamefix } = Utils;
      function getEntry(entry) {
        if (entry && _zip) {
          var item;
          if (typeof entry === "string") item = _zip.getEntry(pth.posix.normalize(entry));
          if (typeof entry === "object" && typeof entry.entryName !== "undefined" && typeof entry.header !== "undefined") item = _zip.getEntry(entry.entryName);
          if (item) {
            return item;
          }
        }
        return null;
      }
      function fixPath(zipPath) {
        const { join: join2, normalize, sep } = pth.posix;
        return join2(".", normalize(sep + zipPath.split("\\").join(sep) + sep));
      }
      function filenameFilter(filterfn) {
        if (filterfn instanceof RegExp) {
          return /* @__PURE__ */ function(rx) {
            return function(filename) {
              return rx.test(filename);
            };
          }(filterfn);
        } else if ("function" !== typeof filterfn) {
          return () => true;
        }
        return filterfn;
      }
      const relativePath = (local, entry) => {
        let lastChar = entry.slice(-1);
        lastChar = lastChar === filetools.sep ? filetools.sep : "";
        return pth.relative(local, entry) + lastChar;
      };
      return {
        /**
         * Extracts the given entry from the archive and returns the content as a Buffer object
         * @param {ZipEntry|string} entry ZipEntry object or String with the full path of the entry
         * @param {Buffer|string} [pass] - password
         * @return Buffer or Null in case of error
         */
        readFile: function(entry, pass) {
          var item = getEntry(entry);
          return item && item.getData(pass) || null;
        },
        /**
         * Returns how many child elements has on entry (directories) on files it is always 0
         * @param {ZipEntry|string} entry ZipEntry object or String with the full path of the entry
         * @returns {integer}
         */
        childCount: function(entry) {
          const item = getEntry(entry);
          if (item) {
            return _zip.getChildCount(item);
          }
        },
        /**
         * Asynchronous readFile
         * @param {ZipEntry|string} entry ZipEntry object or String with the full path of the entry
         * @param {callback} callback
         *
         * @return Buffer or Null in case of error
         */
        readFileAsync: function(entry, callback) {
          var item = getEntry(entry);
          if (item) {
            item.getDataAsync(callback);
          } else {
            callback(null, "getEntry failed for:" + entry);
          }
        },
        /**
         * Extracts the given entry from the archive and returns the content as plain text in the given encoding
         * @param {ZipEntry|string} entry - ZipEntry object or String with the full path of the entry
         * @param {string} encoding - Optional. If no encoding is specified utf8 is used
         *
         * @return String
         */
        readAsText: function(entry, encoding) {
          var item = getEntry(entry);
          if (item) {
            var data = item.getData();
            if (data && data.length) {
              return data.toString(encoding || "utf8");
            }
          }
          return "";
        },
        /**
         * Asynchronous readAsText
         * @param {ZipEntry|string} entry ZipEntry object or String with the full path of the entry
         * @param {callback} callback
         * @param {string} [encoding] - Optional. If no encoding is specified utf8 is used
         *
         * @return String
         */
        readAsTextAsync: function(entry, callback, encoding) {
          var item = getEntry(entry);
          if (item) {
            item.getDataAsync(function(data, err) {
              if (err) {
                callback(data, err);
                return;
              }
              if (data && data.length) {
                callback(data.toString(encoding || "utf8"));
              } else {
                callback("");
              }
            });
          } else {
            callback("");
          }
        },
        /**
         * Remove the entry from the file or the entry and all it's nested directories and files if the given entry is a directory
         *
         * @param {ZipEntry|string} entry
         * @returns {void}
         */
        deleteFile: function(entry, withsubfolders = true) {
          var item = getEntry(entry);
          if (item) {
            _zip.deleteFile(item.entryName, withsubfolders);
          }
        },
        /**
         * Remove the entry from the file or directory without affecting any nested entries
         *
         * @param {ZipEntry|string} entry
         * @returns {void}
         */
        deleteEntry: function(entry) {
          var item = getEntry(entry);
          if (item) {
            _zip.deleteEntry(item.entryName);
          }
        },
        /**
         * Adds a comment to the zip. The zip must be rewritten after adding the comment.
         *
         * @param {string} comment
         */
        addZipComment: function(comment) {
          _zip.comment = comment;
        },
        /**
         * Returns the zip comment
         *
         * @return String
         */
        getZipComment: function() {
          return _zip.comment || "";
        },
        /**
         * Adds a comment to a specified zipEntry. The zip must be rewritten after adding the comment
         * The comment cannot exceed 65535 characters in length
         *
         * @param {ZipEntry} entry
         * @param {string} comment
         */
        addZipEntryComment: function(entry, comment) {
          var item = getEntry(entry);
          if (item) {
            item.comment = comment;
          }
        },
        /**
         * Returns the comment of the specified entry
         *
         * @param {ZipEntry} entry
         * @return String
         */
        getZipEntryComment: function(entry) {
          var item = getEntry(entry);
          if (item) {
            return item.comment || "";
          }
          return "";
        },
        /**
         * Updates the content of an existing entry inside the archive. The zip must be rewritten after updating the content
         *
         * @param {ZipEntry} entry
         * @param {Buffer} content
         */
        updateFile: function(entry, content) {
          var item = getEntry(entry);
          if (item) {
            item.setData(content);
          }
        },
        /**
         * Adds a file from the disk to the archive
         *
         * @param {string} localPath File to add to zip
         * @param {string} [zipPath] Optional path inside the zip
         * @param {string} [zipName] Optional name for the file
         * @param {string} [comment] Optional file comment
         */
        addLocalFile: function(localPath2, zipPath, zipName, comment) {
          if (filetools.fs.existsSync(localPath2)) {
            zipPath = zipPath ? fixPath(zipPath) : "";
            const p = pth.win32.basename(pth.win32.normalize(localPath2));
            zipPath += zipName ? zipName : p;
            const _attr = filetools.fs.statSync(localPath2);
            const data = _attr.isFile() ? filetools.fs.readFileSync(localPath2) : Buffer.alloc(0);
            if (_attr.isDirectory()) zipPath += filetools.sep;
            this.addFile(zipPath, data, comment, _attr);
          } else {
            throw Utils.Errors.FILE_NOT_FOUND(localPath2);
          }
        },
        /**
         * Callback for showing if everything was done.
         *
         * @callback doneCallback
         * @param {Error} err - Error object
         * @param {boolean} done - was request fully completed
         */
        /**
         * Adds a file from the disk to the archive
         *
         * @param {(object|string)} options - options object, if it is string it us used as localPath.
         * @param {string} options.localPath - Local path to the file.
         * @param {string} [options.comment] - Optional file comment.
         * @param {string} [options.zipPath] - Optional path inside the zip
         * @param {string} [options.zipName] - Optional name for the file
         * @param {doneCallback} callback - The callback that handles the response.
         */
        addLocalFileAsync: function(options2, callback) {
          options2 = typeof options2 === "object" ? options2 : { localPath: options2 };
          const localPath2 = pth.resolve(options2.localPath);
          const { comment } = options2;
          let { zipPath, zipName } = options2;
          const self = this;
          filetools.fs.stat(localPath2, function(err, stats) {
            if (err) return callback(err, false);
            zipPath = zipPath ? fixPath(zipPath) : "";
            const p = pth.win32.basename(pth.win32.normalize(localPath2));
            zipPath += zipName ? zipName : p;
            if (stats.isFile()) {
              filetools.fs.readFile(localPath2, function(err2, data) {
                if (err2) return callback(err2, false);
                self.addFile(zipPath, data, comment, stats);
                return setImmediate(callback, void 0, true);
              });
            } else if (stats.isDirectory()) {
              zipPath += filetools.sep;
              self.addFile(zipPath, Buffer.alloc(0), comment, stats);
              return setImmediate(callback, void 0, true);
            }
          });
        },
        /**
         * Adds a local directory and all its nested files and directories to the archive
         *
         * @param {string} localPath - local path to the folder
         * @param {string} [zipPath] - optional path inside zip
         * @param {(RegExp|function)} [filter] - optional RegExp or Function if files match will be included.
         */
        addLocalFolder: function(localPath2, zipPath, filter) {
          filter = filenameFilter(filter);
          zipPath = zipPath ? fixPath(zipPath) : "";
          localPath2 = pth.normalize(localPath2);
          if (filetools.fs.existsSync(localPath2)) {
            const items = filetools.findFiles(localPath2);
            const self = this;
            if (items.length) {
              for (const filepath of items) {
                const p = pth.join(zipPath, relativePath(localPath2, filepath));
                if (filter(p)) {
                  self.addLocalFile(filepath, pth.dirname(p));
                }
              }
            }
          } else {
            throw Utils.Errors.FILE_NOT_FOUND(localPath2);
          }
        },
        /**
         * Asynchronous addLocalFolder
         * @param {string} localPath
         * @param {callback} callback
         * @param {string} [zipPath] optional path inside zip
         * @param {RegExp|function} [filter] optional RegExp or Function if files match will
         *               be included.
         */
        addLocalFolderAsync: function(localPath2, callback, zipPath, filter) {
          filter = filenameFilter(filter);
          zipPath = zipPath ? fixPath(zipPath) : "";
          localPath2 = pth.normalize(localPath2);
          var self = this;
          filetools.fs.open(localPath2, "r", function(err) {
            if (err && err.code === "ENOENT") {
              callback(void 0, Utils.Errors.FILE_NOT_FOUND(localPath2));
            } else if (err) {
              callback(void 0, err);
            } else {
              var items = filetools.findFiles(localPath2);
              var i = -1;
              var next = function() {
                i += 1;
                if (i < items.length) {
                  var filepath = items[i];
                  var p = relativePath(localPath2, filepath).split("\\").join("/");
                  p = p.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^\x20-\x7E]/g, "");
                  if (filter(p)) {
                    filetools.fs.stat(filepath, function(er0, stats) {
                      if (er0) callback(void 0, er0);
                      if (stats.isFile()) {
                        filetools.fs.readFile(filepath, function(er1, data) {
                          if (er1) {
                            callback(void 0, er1);
                          } else {
                            self.addFile(zipPath + p, data, "", stats);
                            next();
                          }
                        });
                      } else {
                        self.addFile(zipPath + p + "/", Buffer.alloc(0), "", stats);
                        next();
                      }
                    });
                  } else {
                    process.nextTick(() => {
                      next();
                    });
                  }
                } else {
                  callback(true, void 0);
                }
              };
              next();
            }
          });
        },
        /**
         * Adds a local directory and all its nested files and directories to the archive
         *
         * @param {object | string} options - options object, if it is string it us used as localPath.
         * @param {string} options.localPath - Local path to the folder.
         * @param {string} [options.zipPath] - optional path inside zip.
         * @param {RegExp|function} [options.filter] - optional RegExp or Function if files match will be included.
         * @param {function|string} [options.namefix] - optional function to help fix filename
         * @param {doneCallback} callback - The callback that handles the response.
         *
         */
        addLocalFolderAsync2: function(options2, callback) {
          const self = this;
          options2 = typeof options2 === "object" ? options2 : { localPath: options2 };
          localPath = pth.resolve(fixPath(options2.localPath));
          let { zipPath, filter, namefix } = options2;
          if (filter instanceof RegExp) {
            filter = /* @__PURE__ */ function(rx) {
              return function(filename) {
                return rx.test(filename);
              };
            }(filter);
          } else if ("function" !== typeof filter) {
            filter = function() {
              return true;
            };
          }
          zipPath = zipPath ? fixPath(zipPath) : "";
          if (namefix == "latin1") {
            namefix = (str) => str.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^\x20-\x7E]/g, "");
          }
          if (typeof namefix !== "function") namefix = (str) => str;
          const relPathFix = (entry) => pth.join(zipPath, namefix(relativePath(localPath, entry)));
          const fileNameFix = (entry) => pth.win32.basename(pth.win32.normalize(namefix(entry)));
          filetools.fs.open(localPath, "r", function(err) {
            if (err && err.code === "ENOENT") {
              callback(void 0, Utils.Errors.FILE_NOT_FOUND(localPath));
            } else if (err) {
              callback(void 0, err);
            } else {
              filetools.findFilesAsync(localPath, function(err2, fileEntries) {
                if (err2) return callback(err2);
                fileEntries = fileEntries.filter((dir) => filter(relPathFix(dir)));
                if (!fileEntries.length) callback(void 0, false);
                setImmediate(
                  fileEntries.reverse().reduce(function(next, entry) {
                    return function(err3, done) {
                      if (err3 || done === false) return setImmediate(next, err3, false);
                      self.addLocalFileAsync(
                        {
                          localPath: entry,
                          zipPath: pth.dirname(relPathFix(entry)),
                          zipName: fileNameFix(entry)
                        },
                        next
                      );
                    };
                  }, callback)
                );
              });
            }
          });
        },
        /**
         * Adds a local directory and all its nested files and directories to the archive
         *
         * @param {string} localPath - path where files will be extracted
         * @param {object} props - optional properties
         * @param {string} [props.zipPath] - optional path inside zip
         * @param {RegExp|function} [props.filter] - optional RegExp or Function if files match will be included.
         * @param {function|string} [props.namefix] - optional function to help fix filename
         */
        addLocalFolderPromise: function(localPath2, props) {
          return new Promise((resolve, reject) => {
            this.addLocalFolderAsync2(Object.assign({ localPath: localPath2 }, props), (err, done) => {
              if (err) reject(err);
              if (done) resolve(this);
            });
          });
        },
        /**
         * Allows you to create a entry (file or directory) in the zip file.
         * If you want to create a directory the entryName must end in / and a null buffer should be provided.
         * Comment and attributes are optional
         *
         * @param {string} entryName
         * @param {Buffer | string} content - file content as buffer or utf8 coded string
         * @param {string} [comment] - file comment
         * @param {number | object} [attr] - number as unix file permissions, object as filesystem Stats object
         */
        addFile: function(entryName, content, comment, attr) {
          entryName = zipnamefix(entryName);
          let entry = getEntry(entryName);
          const update = entry != null;
          if (!update) {
            entry = new ZipEntry(opts);
            entry.entryName = entryName;
          }
          entry.comment = comment || "";
          const isStat = "object" === typeof attr && attr instanceof filetools.fs.Stats;
          if (isStat) {
            entry.header.time = attr.mtime;
          }
          var fileattr = entry.isDirectory ? 16 : 0;
          let unix = entry.isDirectory ? 16384 : 32768;
          if (isStat) {
            unix |= 4095 & attr.mode;
          } else if ("number" === typeof attr) {
            unix |= 4095 & attr;
          } else {
            unix |= entry.isDirectory ? 493 : 420;
          }
          fileattr = (fileattr | unix << 16) >>> 0;
          entry.attr = fileattr;
          entry.setData(content);
          if (!update) _zip.setEntry(entry);
          return entry;
        },
        /**
         * Returns an array of ZipEntry objects representing the files and folders inside the archive
         *
         * @param {string} [password]
         * @returns Array
         */
        getEntries: function(password) {
          _zip.password = password;
          return _zip ? _zip.entries : [];
        },
        /**
         * Returns a ZipEntry object representing the file or folder specified by ``name``.
         *
         * @param {string} name
         * @return ZipEntry
         */
        getEntry: function(name) {
          return getEntry(name);
        },
        getEntryCount: function() {
          return _zip.getEntryCount();
        },
        forEach: function(callback) {
          return _zip.forEach(callback);
        },
        /**
         * Extracts the given entry to the given targetPath
         * If the entry is a directory inside the archive, the entire directory and it's subdirectories will be extracted
         *
         * @param {string|ZipEntry} entry - ZipEntry object or String with the full path of the entry
         * @param {string} targetPath - Target folder where to write the file
         * @param {boolean} [maintainEntryPath=true] - If maintainEntryPath is true and the entry is inside a folder, the entry folder will be created in targetPath as well. Default is TRUE
         * @param {boolean} [overwrite=false] - If the file already exists at the target path, the file will be overwriten if this is true.
         * @param {boolean} [keepOriginalPermission=false] - The file will be set as the permission from the entry if this is true.
         * @param {string} [outFileName] - String If set will override the filename of the extracted file (Only works if the entry is a file)
         *
         * @return Boolean
         */
        extractEntryTo: function(entry, targetPath, maintainEntryPath, overwrite, keepOriginalPermission, outFileName) {
          overwrite = get_Bool(false, overwrite);
          keepOriginalPermission = get_Bool(false, keepOriginalPermission);
          maintainEntryPath = get_Bool(true, maintainEntryPath);
          outFileName = get_Str(keepOriginalPermission, outFileName);
          var item = getEntry(entry);
          if (!item) {
            throw Utils.Errors.NO_ENTRY();
          }
          var entryName = canonical(item.entryName);
          var target = sanitize(targetPath, outFileName && !item.isDirectory ? outFileName : maintainEntryPath ? entryName : pth.basename(entryName));
          if (item.isDirectory) {
            var children = _zip.getEntryChildren(item);
            children.forEach(function(child) {
              if (child.isDirectory) return;
              var content2 = child.getData();
              if (!content2) {
                throw Utils.Errors.CANT_EXTRACT_FILE();
              }
              var name = canonical(child.entryName);
              var childName = sanitize(targetPath, maintainEntryPath ? name : pth.basename(name));
              const fileAttr2 = keepOriginalPermission ? child.header.fileAttr : void 0;
              filetools.writeFileTo(childName, content2, overwrite, fileAttr2);
            });
            return true;
          }
          var content = item.getData(_zip.password);
          if (!content) throw Utils.Errors.CANT_EXTRACT_FILE();
          if (filetools.fs.existsSync(target) && !overwrite) {
            throw Utils.Errors.CANT_OVERRIDE();
          }
          const fileAttr = keepOriginalPermission ? entry.header.fileAttr : void 0;
          filetools.writeFileTo(target, content, overwrite, fileAttr);
          return true;
        },
        /**
         * Test the archive
         * @param {string} [pass]
         */
        test: function(pass) {
          if (!_zip) {
            return false;
          }
          for (var entry in _zip.entries) {
            try {
              if (entry.isDirectory) {
                continue;
              }
              var content = _zip.entries[entry].getData(pass);
              if (!content) {
                return false;
              }
            } catch (err) {
              return false;
            }
          }
          return true;
        },
        /**
         * Extracts the entire archive to the given location
         *
         * @param {string} targetPath Target location
         * @param {boolean} [overwrite=false] If the file already exists at the target path, the file will be overwriten if this is true.
         *                  Default is FALSE
         * @param {boolean} [keepOriginalPermission=false] The file will be set as the permission from the entry if this is true.
         *                  Default is FALSE
         * @param {string|Buffer} [pass] password
         */
        extractAllTo: function(targetPath, overwrite, keepOriginalPermission, pass) {
          keepOriginalPermission = get_Bool(false, keepOriginalPermission);
          pass = get_Str(keepOriginalPermission, pass);
          overwrite = get_Bool(false, overwrite);
          if (!_zip) throw Utils.Errors.NO_ZIP();
          _zip.entries.forEach(function(entry) {
            var entryName = sanitize(targetPath, canonical(entry.entryName));
            if (entry.isDirectory) {
              filetools.makeDir(entryName);
              return;
            }
            var content = entry.getData(pass);
            if (!content) {
              throw Utils.Errors.CANT_EXTRACT_FILE();
            }
            const fileAttr = keepOriginalPermission ? entry.header.fileAttr : void 0;
            filetools.writeFileTo(entryName, content, overwrite, fileAttr);
            try {
              filetools.fs.utimesSync(entryName, entry.header.time, entry.header.time);
            } catch (err) {
              throw Utils.Errors.CANT_EXTRACT_FILE();
            }
          });
        },
        /**
         * Asynchronous extractAllTo
         *
         * @param {string} targetPath Target location
         * @param {boolean} [overwrite=false] If the file already exists at the target path, the file will be overwriten if this is true.
         *                  Default is FALSE
         * @param {boolean} [keepOriginalPermission=false] The file will be set as the permission from the entry if this is true.
         *                  Default is FALSE
         * @param {function} callback The callback will be executed when all entries are extracted successfully or any error is thrown.
         */
        extractAllToAsync: function(targetPath, overwrite, keepOriginalPermission, callback) {
          callback = get_Fun(overwrite, keepOriginalPermission, callback);
          keepOriginalPermission = get_Bool(false, keepOriginalPermission);
          overwrite = get_Bool(false, overwrite);
          if (!callback) {
            return new Promise((resolve, reject) => {
              this.extractAllToAsync(targetPath, overwrite, keepOriginalPermission, function(err) {
                if (err) {
                  reject(err);
                } else {
                  resolve(this);
                }
              });
            });
          }
          if (!_zip) {
            callback(Utils.Errors.NO_ZIP());
            return;
          }
          targetPath = pth.resolve(targetPath);
          const getPath = (entry) => sanitize(targetPath, pth.normalize(canonical(entry.entryName)));
          const getError = (msg, file) => new Error(msg + ': "' + file + '"');
          const dirEntries = [];
          const fileEntries = [];
          _zip.entries.forEach((e) => {
            if (e.isDirectory) {
              dirEntries.push(e);
            } else {
              fileEntries.push(e);
            }
          });
          for (const entry of dirEntries) {
            const dirPath = getPath(entry);
            const dirAttr = keepOriginalPermission ? entry.header.fileAttr : void 0;
            try {
              filetools.makeDir(dirPath);
              if (dirAttr) filetools.fs.chmodSync(dirPath, dirAttr);
              filetools.fs.utimesSync(dirPath, entry.header.time, entry.header.time);
            } catch (er) {
              callback(getError("Unable to create folder", dirPath));
            }
          }
          fileEntries.reverse().reduce(function(next, entry) {
            return function(err) {
              if (err) {
                next(err);
              } else {
                const entryName = pth.normalize(canonical(entry.entryName));
                const filePath = sanitize(targetPath, entryName);
                entry.getDataAsync(function(content, err_1) {
                  if (err_1) {
                    next(err_1);
                  } else if (!content) {
                    next(Utils.Errors.CANT_EXTRACT_FILE());
                  } else {
                    const fileAttr = keepOriginalPermission ? entry.header.fileAttr : void 0;
                    filetools.writeFileToAsync(filePath, content, overwrite, fileAttr, function(succ) {
                      if (!succ) {
                        next(getError("Unable to write file", filePath));
                      }
                      filetools.fs.utimes(filePath, entry.header.time, entry.header.time, function(err_2) {
                        if (err_2) {
                          next(getError("Unable to set times", filePath));
                        } else {
                          next();
                        }
                      });
                    });
                  }
                });
              }
            };
          }, callback)();
        },
        /**
         * Writes the newly created zip file to disk at the specified location or if a zip was opened and no ``targetFileName`` is provided, it will overwrite the opened zip
         *
         * @param {string} targetFileName
         * @param {function} callback
         */
        writeZip: function(targetFileName, callback) {
          if (arguments.length === 1) {
            if (typeof targetFileName === "function") {
              callback = targetFileName;
              targetFileName = "";
            }
          }
          if (!targetFileName && opts.filename) {
            targetFileName = opts.filename;
          }
          if (!targetFileName) return;
          var zipData = _zip.compressToBuffer();
          if (zipData) {
            var ok = filetools.writeFileTo(targetFileName, zipData, true);
            if (typeof callback === "function") callback(!ok ? new Error("failed") : null, "");
          }
        },
        /**
                 *
                 * @param {string} targetFileName
                 * @param {object} [props]
                 * @param {boolean} [props.overwrite=true] If the file already exists at the target path, the file will be overwriten if this is true.
                 * @param {boolean} [props.perm] The file will be set as the permission from the entry if this is true.
        
                 * @returns {Promise<void>}
                 */
        writeZipPromise: function(targetFileName, props) {
          const { overwrite, perm } = Object.assign({ overwrite: true }, props);
          return new Promise((resolve, reject) => {
            if (!targetFileName && opts.filename) targetFileName = opts.filename;
            if (!targetFileName) reject("ADM-ZIP: ZIP File Name Missing");
            this.toBufferPromise().then((zipData) => {
              const ret = (done) => done ? resolve(done) : reject("ADM-ZIP: Wasn't able to write zip file");
              filetools.writeFileToAsync(targetFileName, zipData, overwrite, perm, ret);
            }, reject);
          });
        },
        /**
         * @returns {Promise<Buffer>} A promise to the Buffer.
         */
        toBufferPromise: function() {
          return new Promise((resolve, reject) => {
            _zip.toAsyncBuffer(resolve, reject);
          });
        },
        /**
         * Returns the content of the entire zip file as a Buffer object
         *
         * @prop {function} [onSuccess]
         * @prop {function} [onFail]
         * @prop {function} [onItemStart]
         * @prop {function} [onItemEnd]
         * @returns {Buffer}
         */
        toBuffer: function(onSuccess, onFail, onItemStart, onItemEnd) {
          if (typeof onSuccess === "function") {
            _zip.toAsyncBuffer(onSuccess, onFail, onItemStart, onItemEnd);
            return null;
          }
          return _zip.compressToBuffer();
        }
      };
    };
  }
});

// main.ts
var main_exports = {};
__export(main_exports, {
  default: () => MD2PDFPlugin
});
module.exports = __toCommonJS(main_exports);
var import_obsidian = require("obsidian");
var import_child_process = require("child_process");
var import_util = require("util");
var import_url = require("url");
var path = __toESM(require("path"));
var fs = __toESM(require("fs"));
var os = __toESM(require("os"));
var execAsync = (0, import_util.promisify)(import_child_process.exec);
var DEFAULT_SETTINGS = {
  pandocPath: "",
  libreOfficePath: "",
  typstPath: "",
  pdfEngine: "typst",
  defaultRef: "",
  deleteIntermediateDocx: true,
  autoDetectBibliography: true,
  customPandocArgs: "",
  pdfOutputLocation: "",
  docxOutputLocation: "",
  offerToOpen: true
};
var FRONTMATTER_PROPS = ["title", "subtitle", "author", "instructor", "course", "date", "spacing", "style"];
var TYPST_SPACINGS = ["single", "double"];
var DEFAULT_TEMPLATE = "folio.typ";
var TEMPLATES = {
  mla: "mla.typ",
  chicago: "turabian.typ",
  turabian: "turabian.typ"
};
var DEFAULT_CSLS = {
  chicago: "chicago-notes-bibliography.csl",
  turabian: "chicago-notes-bibliography.csl",
  mla: "modern-language-association.csl"
};
var MONTH_NAMES = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December"
];
var COMMAND_TIMEOUT_MS = 10 * 60 * 1e3;
var CommandError = class extends Error {
  constructor(stage, args, code, stdout, stderr) {
    super(`${stage} failed${code !== null ? ` (exit ${code})` : ""}: ${lastLine(stderr || stdout) || "no output"}`);
    this.stage = stage;
    this.args = args;
    this.code = code;
    this.stdout = stdout;
    this.stderr = stderr;
  }
};
function lastLine(s) {
  const lines = s.trim().split("\n");
  return lines[lines.length - 1].slice(0, 120);
}
function styleKey(yaml) {
  return (yaml["style"] || "").trim().toLowerCase();
}
function isChicago(yaml) {
  const s = styleKey(yaml);
  return s === "chicago" || s === "turabian";
}
var MD2PDFPlugin = class extends import_obsidian.Plugin {
  async onload() {
    await this.loadSettings();
    this.addCommand({
      id: "convert-md-to-pdf",
      name: "Export current note to PDF",
      editorCallback: async (editor, view) => {
        if (view.file) {
          await this.exportNote(view.file, "pdf");
        } else {
          new import_obsidian.Notice("No file is currently open");
        }
      }
    });
    this.addCommand({
      id: "convert-md-to-docx",
      name: "Export current note to DOCX",
      editorCallback: async (editor, view) => {
        if (view.file) {
          await this.exportNote(view.file, "docx");
        } else {
          new import_obsidian.Notice("No file is currently open");
        }
      }
    });
    this.addCommand({
      id: "insert-frontmatter",
      name: "Insert export frontmatter",
      editorCallback: (editor) => this.insertFrontmatter(editor)
    });
    this.addRibbonIcon("file-output", "Export note to PDF", async () => {
      const activeFile = this.app.workspace.getActiveFile();
      if (activeFile && activeFile.extension === "md") {
        await this.exportNote(activeFile, "pdf");
      } else {
        new import_obsidian.Notice("Please open a Markdown file first");
      }
    });
    this.addSettingTab(new MD2PDFSettingTab(this.app, this));
    if (!this.settings.pandocPath || !this.settings.libreOfficePath || !this.settings.typstPath) {
      await this.autoDetectPaths();
    }
  }
  insertFrontmatter(editor) {
    const text = editor.getValue();
    const m = text.match(/^---\n([\s\S]*?)\n---/);
    if (m) {
      const existing = /* @__PURE__ */ new Set();
      for (const line of m[1].split("\n")) {
        const idx = line.indexOf(":");
        if (idx > 0) existing.add(line.substring(0, idx).trim());
      }
      const missing = FRONTMATTER_PROPS.filter((p) => !existing.has(p));
      if (missing.length === 0) {
        new import_obsidian.Notice("All frontmatter properties already present.");
        return;
      }
      const endOfBlock = editor.offsetToPos(4 + m[1].length);
      editor.replaceRange("\n" + missing.map((p) => `${p}: `).join("\n"), endOfBlock);
    } else {
      const block = FRONTMATTER_PROPS.map((p) => `${p}: `).join("\n");
      editor.replaceRange(`---
${block}
---
`, { line: 0, ch: 0 });
    }
    new import_obsidian.Notice("Frontmatter inserted.");
  }
  async autoDetectPaths() {
    const platform2 = os.platform();
    let pandocPaths = [];
    let libreofficePaths = [];
    let typstPaths = [];
    if (platform2 === "win32") {
      pandocPaths = [
        path.join("C:", "Program Files", "Pandoc", "pandoc.exe"),
        path.join("C:", "Program Files (x86)", "Pandoc", "pandoc.exe"),
        path.join(os.homedir(), "AppData", "Local", "Pandoc", "pandoc.exe"),
        path.join(os.homedir(), "scoop", "shims", "pandoc.exe"),
        path.join("C:", "ProgramData", "chocolatey", "bin", "pandoc.exe")
      ];
      libreofficePaths = [
        path.join("C:", "Program Files", "LibreOffice", "program", "soffice.exe"),
        path.join("C:", "Program Files (x86)", "LibreOffice", "program", "soffice.exe")
      ];
      typstPaths = [
        path.join(os.homedir(), "scoop", "shims", "typst.exe"),
        path.join(os.homedir(), ".cargo", "bin", "typst.exe"),
        path.join(os.homedir(), "AppData", "Local", "Microsoft", "WinGet", "Links", "typst.exe"),
        path.join("C:", "ProgramData", "chocolatey", "bin", "typst.exe")
      ];
      const programFiles = process.env["ProgramFiles"] || path.join("C:", "Program Files");
      const programFilesX86 = process.env["ProgramFiles(x86)"] || path.join("C:", "Program Files (x86)");
      for (const base of [programFiles, programFilesX86]) {
        try {
          const entries = fs.readdirSync(base);
          for (const entry of entries) {
            if (entry.startsWith("LibreOffice")) {
              const sofficePath = path.join(base, entry, "program", "soffice.exe");
              if (!libreofficePaths.includes(sofficePath)) {
                libreofficePaths.push(sofficePath);
              }
            }
          }
        } catch (e) {
        }
      }
    } else if (platform2 === "darwin") {
      pandocPaths = [
        "/usr/local/bin/pandoc",
        "/opt/homebrew/bin/pandoc",
        "/usr/bin/pandoc"
      ];
      libreofficePaths = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/local/bin/soffice"
      ];
      typstPaths = [
        "/opt/homebrew/bin/typst",
        "/usr/local/bin/typst",
        path.join(os.homedir(), ".cargo", "bin", "typst")
      ];
    } else {
      pandocPaths = [
        "/usr/bin/pandoc",
        "/usr/local/bin/pandoc",
        "/snap/bin/pandoc"
      ];
      libreofficePaths = [
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
        "/usr/local/bin/libreoffice",
        "/snap/bin/libreoffice"
      ];
      typstPaths = [
        "/usr/bin/typst",
        "/usr/local/bin/typst",
        "/snap/bin/typst",
        path.join(os.homedir(), ".cargo", "bin", "typst")
      ];
    }
    if (!this.settings.pandocPath) {
      this.settings.pandocPath = await this.findBinary(
        pandocPaths,
        platform2 === "win32" ? "where pandoc" : "which pandoc"
      );
    }
    if (!this.settings.libreOfficePath) {
      this.settings.libreOfficePath = await this.findBinary(
        libreofficePaths,
        platform2 === "win32" ? "where soffice" : "which libreoffice || which soffice"
      );
    }
    if (!this.settings.typstPath) {
      this.settings.typstPath = await this.findBinary(
        typstPaths,
        platform2 === "win32" ? "where typst" : "which typst"
      );
    }
    await this.saveSettings();
  }
  async findBinary(candidates, whichCmd) {
    for (const p of candidates) {
      if (fs.existsSync(p)) return p;
    }
    try {
      const { stdout } = await execAsync(whichCmd);
      const foundPath = stdout.trim().split("\n")[0].trim();
      if (foundPath && fs.existsSync(foundPath)) return foundPath;
    } catch (e) {
    }
    return "";
  }
  async loadSettings() {
    const data = await this.loadData();
    this.settings = Object.assign({}, DEFAULT_SETTINGS, data);
    if (data && !data.pdfEngine) {
      this.settings.pdfEngine = "libreoffice";
    }
  }
  async saveSettings() {
    await this.saveData(this.settings);
  }
  getVaultPath() {
    return this.app.vault.adapter.basePath;
  }
  // Get plugin folder path
  getPluginFolder() {
    const vaultPath = this.getVaultPath();
    const pluginDir = this.manifest.dir || path.join(this.app.vault.configDir, "plugins", this.manifest.id);
    return path.join(vaultPath, pluginDir);
  }
  // Get list of available reference.docx files in plugin folder
  getAvailableRefDocs() {
    const pluginFolder = this.getPluginFolder();
    try {
      if (fs.existsSync(pluginFolder)) {
        const files = fs.readdirSync(pluginFolder);
        return files.filter((f) => f.endsWith(".docx")).map((f) => f.replace(".docx", ""));
      }
    } catch (e) {
      console.warn("Could not read plugin folder:", e);
    }
    return [];
  }
  // Parse YAML frontmatter from content. Deliberately splits on the first
  // colon rather than parsing real YAML: academic titles are full of
  // unquoted colons, which strict YAML rejects.
  parseYAML(content) {
    const yaml = {};
    const match = content.match(/^---\r?\n([\s\S]*?)\r?\n---/);
    if (match) {
      const lines = match[1].split(/\r?\n/);
      for (const line of lines) {
        const colonIndex = line.indexOf(":");
        if (colonIndex > 0) {
          const key = line.substring(0, colonIndex).trim();
          let value = line.substring(colonIndex + 1).trim();
          if (value.length >= 2 && value[0] === value[value.length - 1] && (value[0] === '"' || value[0] === "'")) {
            value = value.slice(1, -1);
          }
          yaml[key] = value;
        }
      }
    }
    return yaml;
  }
  resolveReferenceDoc(yaml) {
    const pluginFolder = this.getPluginFolder();
    if (yaml["spacing"]) {
      const refPath = path.join(pluginFolder, yaml["spacing"] + ".docx");
      if (fs.existsSync(refPath)) return refPath;
      new import_obsidian.Notice(`Reference doc not found: ${refPath}`);
      return null;
    }
    if (this.settings.defaultRef) {
      const defaultRefPath = path.join(pluginFolder, this.settings.defaultRef + ".docx");
      if (fs.existsSync(defaultRefPath)) return defaultRefPath;
    }
    const availableDocs = this.getAvailableRefDocs();
    if (availableDocs.includes("double")) return path.join(pluginFolder, "double.docx");
    if (availableDocs.length > 0) return path.join(pluginFolder, availableDocs[0] + ".docx");
    new import_obsidian.Notice(`No .docx reference document found in ${pluginFolder}`);
    console.error("md2pdf: no reference .docx in plugin folder", {
      pluginFolder,
      manifestDir: this.manifest.dir,
      manifestId: this.manifest.id,
      folderExists: fs.existsSync(pluginFolder)
    });
    return null;
  }
  templatePath(yaml) {
    return path.join(this.getPluginFolder(), "templates", TEMPLATES[styleKey(yaml)] || DEFAULT_TEMPLATE);
  }
  // Which engine renders PDF for this note. A spacing: the Typst templates
  // cannot express names a reference .docx, so only LibreOffice can honour
  // it; otherwise the setting decides, and Typst is only used when both the
  // binary and the template are actually there.
  pdfEngineFor(yaml) {
    const spacing = yaml["spacing"];
    if (spacing && !TYPST_SPACINGS.includes(spacing)) return "libreoffice";
    if (this.settings.pdfEngine === "libreoffice") return "libreoffice";
    const typstOk = !!this.settings.typstPath && fs.existsSync(this.settings.typstPath) && fs.existsSync(this.templatePath(yaml));
    return typstOk ? "typst" : "libreoffice";
  }
  resolveOutputDir(setting, fallback) {
    if (!setting) return fallback;
    const dir = path.isAbsolute(setting) ? setting : path.join(this.getVaultPath(), setting);
    fs.mkdirSync(dir, { recursive: true });
    return dir;
  }
  // Locate the .bib named by bibliography:. Tries the literal value
  // (absolute, then vault-relative), then the vault's own .bib — so a bare
  // `bibliography: true`, or an absolute path written on another machine,
  // still finds it.
  resolveBibPath(yaml) {
    const vaultPath = this.getVaultPath();
    const val = (yaml["bibliography"] || "").trim();
    if (val) {
      for (const cand of [expandHome(val), path.join(vaultPath, val)]) {
        if (isFile(cand)) return cand;
      }
    }
    return findBibFile(vaultPath);
  }
  // Locate the .csl for this note. An explicit csl: always wins; only when
  // the note names none (or names one that is not on this machine) does the
  // style: default apply.
  resolveCslPath(yaml) {
    const vaultPath = this.getVaultPath();
    const val = (yaml["csl"] || "").trim();
    if (val) {
      const name = path.basename(val);
      for (const cand of [
        expandHome(val),
        path.join(vaultPath, val),
        path.join(vaultPath, "sources", name),
        path.join(vaultPath, name)
      ]) {
        if (isFile(cand)) return cand;
      }
    }
    const def = DEFAULT_CSLS[styleKey(yaml)];
    if (!def) return null;
    const p = path.join(this.getPluginFolder(), "csl", def);
    return isFile(p) ? p : null;
  }
  async exportNote(file, format) {
    const content = await this.app.vault.read(file);
    const yaml = this.parseYAML(content);
    const engine = format === "pdf" ? this.pdfEngineFor(yaml) : "libreoffice";
    if (!this.settings.pandocPath || !fs.existsSync(this.settings.pandocPath)) {
      new import_obsidian.Notice(this.settings.pandocPath ? `Pandoc not found at: ${this.settings.pandocPath}` : "Pandoc path not configured. Please set it in plugin settings.");
      return;
    }
    if (format === "pdf" && engine === "libreoffice" && (!this.settings.libreOfficePath || !fs.existsSync(this.settings.libreOfficePath))) {
      new import_obsidian.Notice(this.settings.libreOfficePath ? `LibreOffice not found at: ${this.settings.libreOfficePath}` : "LibreOffice path not configured. Install Typst or LibreOffice, or set the path in plugin settings.");
      return;
    }
    let referenceDoc = null;
    if (engine === "libreoffice") {
      referenceDoc = this.resolveReferenceDoc(yaml);
      if (!referenceDoc) return;
    }
    const wantsBib = this.settings.autoDetectBibliography && "bibliography" in yaml;
    const bibSrc = wantsBib ? this.resolveBibPath(yaml) : null;
    if (wantsBib && !bibSrc) {
      new import_obsidian.Notice(`bibliography: not found \u2014 ${yaml["bibliography"]}`);
      return;
    }
    const cslSrc = bibSrc ? this.resolveCslPath(yaml) : null;
    const vaultPath = this.getVaultPath();
    const inputPath = path.join(vaultPath, file.path);
    const noteDir = path.dirname(inputPath);
    const baseName = path.basename(file.path, ".md");
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "md2pdf-"));
    const progress = new import_obsidian.Notice("", 0);
    const step = (msg) => progress.setMessage(`Exporting ${baseName}\u2026 ${msg}`);
    try {
      const pdfOutputDir = this.resolveOutputDir(this.settings.pdfOutputLocation, noteDir);
      const docxOutputDir = this.resolveOutputDir(this.settings.docxOutputLocation, noteDir);
      const pdfPath = path.join(pdfOutputDir, `${baseName}.pdf`);
      const keepDocx = format === "docx" || !this.settings.deleteIntermediateDocx;
      const docxPath = path.join(keepDocx ? docxOutputDir : tmpDir, `${baseName}.docx`);
      const commonArgs = [`--resource-path=${[noteDir, vaultPath].join(path.delimiter)}`];
      if (bibSrc) {
        const bibUse = narrowBib(bibSrc, content, tmpDir);
        commonArgs.push("--citeproc", `--bibliography=${bibUse}`);
        if (cslSrc) commonArgs.push(`--csl=${cslSrc}`);
      }
      const customArgs = this.settings.customPandocArgs.split(/\s+/).filter((arg) => arg.length > 0);
      if (engine === "typst") {
        const bodyMd = path.join(tmpDir, "body.md");
        fs.writeFileSync(bodyMd, wrapBibliography(stripFrontmatter(content)), "utf8");
        step(bibSrc ? "(1/2) Running pandoc with citations" : "(1/2) Running pandoc");
        const pandocArgs = [bodyMd, ...commonArgs];
        if (!bibSrc) pandocArgs.push("--from", "markdown-citations");
        pandocArgs.push(
          "--extract-media=media",
          ...customArgs,
          "--to",
          "typst-citations",
          "-o",
          "body.typ"
        );
        await this.runCommand("pandoc", this.settings.pandocPath, pandocArgs, tmpDir);
        fs.copyFileSync(this.templatePath(yaml), path.join(tmpDir, "folio.typ"));
        fs.writeFileSync(path.join(tmpDir, "main.typ"), typstWrapper(yaml, "body.typ"), "utf8");
        step("(2/2) Rendering PDF");
        const typstArgs = ["compile", "--root", tmpDir];
        const fontsDir = path.join(this.getPluginFolder(), "fonts");
        if (fs.existsSync(fontsDir)) typstArgs.push("--font-path", fontsDir);
        typstArgs.push(path.join(tmpDir, "main.typ"), pdfPath);
        await this.runCommand("typst", this.settings.typstPath, typstArgs, tmpDir);
        if (!fs.existsSync(pdfPath)) throw new Error("typst ran but produced no PDF");
        progress.hide();
        this.exported(pdfPath);
        return;
      }
      const sourceMd = path.join(tmpDir, "source.md");
      fs.writeFileSync(sourceMd, wrapBibliography(stripFrontmatter(content)), "utf8");
      const metaArgs = [];
      if (styleKey(yaml) !== "mla" && !isChicago(yaml)) {
        for (const k of ["title", "subtitle", "author", "date"]) {
          if (yaml[k]) metaArgs.push(`--metadata=${k}:${yaml[k]}`);
        }
      }
      const luaPath = path.join(tmpDir, "filter.lua");
      fs.writeFileSync(luaPath, this.generateLuaFilter(yaml), "utf8");
      const total = format === "pdf" ? 3 : 2;
      step(`(1/${total}) Running pandoc${bibSrc ? " with citations" : ""}`);
      await this.runCommand("pandoc", this.settings.pandocPath, [
        sourceMd,
        "--standalone",
        "--reference-doc=" + referenceDoc,
        "--lua-filter=" + luaPath,
        ...metaArgs,
        ...commonArgs,
        ...customArgs,
        "-o",
        docxPath
      ], noteDir);
      if (!fs.existsSync(docxPath)) throw new Error("Pandoc failed to create DOCX file");
      step(`(2/${total}) Post-processing`);
      this.postprocessDocx(docxPath, yaml);
      if (format === "docx") {
        progress.hide();
        this.exported(docxPath);
        return;
      }
      step("(3/3) Converting to PDF");
      const loProfile = (0, import_url.pathToFileURL)(path.join(tmpDir, "loprofile")).href;
      await this.runCommand("LibreOffice", this.settings.libreOfficePath, [
        `-env:UserInstallation=${loProfile}`,
        "--headless",
        "--norestore",
        "--convert-to",
        "pdf",
        "--outdir",
        pdfOutputDir,
        docxPath
      ], pdfOutputDir);
      if (!fs.existsSync(pdfPath)) throw new Error("LibreOffice ran but produced no PDF");
      progress.hide();
      this.exported(pdfPath);
    } catch (error) {
      progress.hide();
      console.error("MD2PDF export error:", error);
      let msg = `Export failed: ${error.message}`;
      if (error instanceof CommandError) {
        const log = this.logExportError(error, file.path, format);
        if (log) msg += ` (see ${log})`;
      }
      new import_obsidian.Notice(msg, 15e3);
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  }
  exported(outPath) {
    new import_obsidian.Notice(`Exported: ${path.basename(outPath)}`);
    if (this.settings.offerToOpen) {
      new AfterExportModal(this.app, outPath).open();
    }
  }
  // Notices are fleeting and truncated, so write the full command and output
  // somewhere durable. A dot-folder keeps it out of the file explorer.
  logExportError(err, note, format) {
    try {
      const rel = path.join(".md2pdf", "export-error.log");
      const logPath = path.join(this.getVaultPath(), rel);
      fs.mkdirSync(path.dirname(logPath), { recursive: true });
      fs.writeFileSync(
        logPath,
        `md2pdf export failed
stage:   ${err.stage}
format:  ${format}
note:    ${note}
command:
  ${err.args.join(" ")}
returncode: ${err.code}
--- stdout ---
${err.stdout}
--- stderr ---
${err.stderr}
`,
        "utf8"
      );
      return rel;
    } catch (e) {
      return null;
    }
  }
  // One pass over the docx: strip headers or footers by style, fill in the
  // {{LASTNAME}} placeholder, and re-size table columns to their contents.
  // style: mla -> keep headers, strip footers
  // style: chicago or blank -> keep footers, strip headers
  postprocessDocx(docxPath, yaml) {
    const AdmZip = require_adm_zip();
    const isMla = styleKey(yaml) === "mla";
    const lastName = authorLastname(yaml);
    try {
      const zip = new AdmZip(docxPath);
      for (const entry of zip.getEntries()) {
        const name = entry.entryName;
        const isHeader = /^word\/header\d*\.xml$/.test(name);
        const isFooter = /^word\/footer\d*\.xml$/.test(name);
        let content = null;
        if (name === "word/document.xml") {
          content = autofitTables(entry.getData().toString("utf8"));
        } else if (isHeader && !isMla) {
          content = EMPTY_HEADER_XML;
        } else if (isFooter && isMla) {
          content = EMPTY_FOOTER_XML;
        } else if (isHeader || isFooter) {
          const text = entry.getData().toString("utf8");
          content = lastName ? text.split("{{LASTNAME}}").join(lastName) : text.split("{{LASTNAME}} ").join("").split("{{LASTNAME}}").join("");
        }
        if (content !== null) zip.updateFile(name, Buffer.from(content, "utf8"));
      }
      zip.writeZip(docxPath);
    } catch (e) {
      console.warn("Failed to post-process docx:", e);
    }
  }
  generateLuaFilter(yaml) {
    const s = styleKey(yaml);
    if (s === "mla") return this.generateHeaderFilter(yaml);
    if (isChicago(yaml)) return this.generateCoverPageFilter(yaml);
    return this.generateBasicFilter();
  }
  luaMetaVars(yaml) {
    const keys = ["title", "subtitle", "author", "course", "instructor", "date"];
    return `local M = {
${keys.map((k) => `  ${k} = ${luaStr(yaml[k] || "")},`).join("\n")}
}

-- Capture from document metadata when not in the frontmatter we parsed,
-- then remove so pandoc does not also render them in its own styles.
function Meta(meta)
  for k, v in pairs(M) do
    if meta[k] and v == "" then
      M[k] = pandoc.utils.stringify(meta[k])
    end
    meta[k] = nil
  end
  return meta
end`;
  }
  luaBibHelpers() {
    return `local function escape_xml(s)
  s = s:gsub("&", "&amp;")
  s = s:gsub("<", "&lt;")
  s = s:gsub(">", "&gt;")
  return s
end

-- A plain double-spaced paragraph: before is spacing in twips, center a bool.
local function para_xml(text, before, center)
  return pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:spacing w:before="%d" w:after="0" w:line="480" w:lineRule="auto"/>%s
  </w:pPr>
  <w:r>
    <w:t xml:space="preserve">%s</w:t>
  </w:r>
</w:p>]], before or 0, center and '\\n    <w:jc w:val="center"/>' or '', escape_xml(text)))
end

local function format_date(date_str, mla)
  local months = {
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  }
  local year, month, day = date_str:match("(%d+)-(%d+)-(%d+)")
  if year and month and day then
    local month_name = months[tonumber(month)]
    if month_name then
      if mla then
        return string.format("%d %s %s", tonumber(day), month_name, year)
      end
      return string.format("%s %d, %s", month_name, tonumber(day), year)
    end
  end
  return date_str
end

local function inlines_to_openxml(inlines)
  local runs = {}
  for _, inl in ipairs(inlines) do
    if inl.t == "Emph" then
      local txt = escape_xml(pandoc.utils.stringify(inl))
      table.insert(runs, string.format(
        '<w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t xml:space="preserve">%s</w:t></w:r>', txt))
    elseif inl.t == "Strong" then
      local txt = escape_xml(pandoc.utils.stringify(inl))
      table.insert(runs, string.format(
        '<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">%s</w:t></w:r>', txt))
    elseif inl.t == "Str" then
      table.insert(runs, string.format(
        '<w:r><w:t xml:space="preserve">%s</w:t></w:r>', escape_xml(inl.text)))
    elseif inl.t == "Space" then
      table.insert(runs, '<w:r><w:t xml:space="preserve"> </w:t></w:r>')
    elseif inl.t == "SoftBreak" or inl.t == "LineBreak" then
      table.insert(runs, '<w:r><w:t xml:space="preserve"> </w:t></w:r>')
    elseif inl.t == "Link" then
      local txt = escape_xml(pandoc.utils.stringify(inl))
      table.insert(runs, string.format(
        '<w:r><w:t xml:space="preserve">%s</w:t></w:r>', txt))
    else
      local txt = escape_xml(pandoc.utils.stringify(inl))
      if txt ~= "" then
        table.insert(runs, string.format(
          '<w:r><w:t xml:space="preserve">%s</w:t></w:r>', txt))
      end
    end
  end
  return table.concat(runs)
end

local function bib_entry_block(block)
  local runs_xml = inlines_to_openxml(block.content)
  return pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:spacing w:after="0" w:line="480" w:lineRule="auto"/>
    <w:ind w:left="720" w:hanging="720"/>
  </w:pPr>
  %s
</w:p>]], runs_xml))
end

local function is_bib_heading(block)
  if block.t ~= "Header" then return false end
  local text = pandoc.utils.stringify(block):lower()
  return text:match("bibliography") or text:match("references") or text:match("works cited")
end

local function is_refs_div(block)
  if block.t ~= "Div" then return false end
  if block.identifier == "bibentries" then return true end
  for _, cls in ipairs(block.classes) do
    if cls == "references" or cls == "csl-bib-body" then return true end
  end
  return false
end

-- Returns true if the block was consumed as part of a bibliography.
local bib_state = { in_bib = false }
local function handle_bib_block(block, new_blocks)
  if is_bib_heading(block) then
    bib_state.in_bib = true
    table.insert(new_blocks, pandoc.RawBlock('openxml', string.format([[
<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading%d"/>
    <w:pageBreakBefore/>
  </w:pPr>
  <w:r>
    <w:t xml:space="preserve">%s</w:t>
  </w:r>
</w:p>]], block.level, escape_xml(pandoc.utils.stringify(block)))))
    return true
  elseif bib_state.in_bib and block.t == "Header" then
    bib_state.in_bib = false
    return false
  elseif bib_state.in_bib and block.t == "Para" then
    table.insert(new_blocks, bib_entry_block(block))
    return true
  elseif (bib_state.in_bib or is_refs_div(block)) and block.t == "Div" then
    for _, inner in ipairs(block.content) do
      if inner.t == "Div" then
        for _, b in ipairs(inner.content) do
          if b.t == "Para" then
            table.insert(new_blocks, bib_entry_block(b))
          end
        end
      elseif inner.t == "Para" then
        table.insert(new_blocks, bib_entry_block(inner))
      end
    end
    return true
  end
  return false
end`;
  }
  generateBasicFilter() {
    return `${this.luaBibHelpers()}

function Pandoc(doc)
  local new_blocks = {}
  for _, block in ipairs(doc.blocks) do
    if not handle_bib_block(block, new_blocks) then
      table.insert(new_blocks, block)
    end
  end
  doc.blocks = new_blocks
  return doc
end`;
  }
  // Turabian cover page: title centered ~1/3 down, author/course/instructor/
  // date in the bottom half, body from page two.
  generateCoverPageFilter(yaml) {
    return `${this.luaBibHelpers()}
${this.luaMetaVars(yaml)}

function Pandoc(doc)
  local new_blocks = {}

  if M.title ~= "" then
    -- With a subtitle, the title line ends with a colon and the subtitle
    -- takes the line below. A title typed with its own colon keeps just one.
    local title_line = M.title
    if M.subtitle ~= "" and not M.title:match(":%s*$") then
      title_line = M.title .. ":"
    end
    table.insert(new_blocks, para_xml(title_line, 2400, true))
    if M.subtitle ~= "" then
      table.insert(new_blocks, para_xml(M.subtitle, 0, true))
    end
  end

  -- Roughly the bottom third of the page
  local gap = 4320
  local info = { M.author, M.course, M.instructor,
                 M.date ~= "" and format_date(M.date, false) or "" }
  for _, line in ipairs(info) do
    if line ~= "" then
      table.insert(new_blocks, para_xml(line, gap, true))
      gap = 0
    end
  end

  local page_break_inserted = false
  for _, block in ipairs(doc.blocks) do
    if is_bib_heading(block) then
      page_break_inserted = true
    end
    if not handle_bib_block(block, new_blocks) then
      if not page_break_inserted and (block.t == "Header" or
         (block.t == "Para" and #block.content > 0) or
         block.t == "CodeBlock" or block.t == "BulletList" or
         block.t == "OrderedList" or block.t == "Table" or
         block.t == "BlockQuote" or block.t == "RawBlock" or block.t == "Div") then
        table.insert(new_blocks, pandoc.RawBlock('openxml',
          '<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>'))
        page_break_inserted = true
      end
      table.insert(new_blocks, block)
    end
  end

  if not page_break_inserted then
    table.insert(new_blocks, pandoc.RawBlock('openxml',
      '<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>'))
  end

  doc.blocks = new_blocks
  return doc
end`;
  }
  // MLA: author, instructor, course, date left-aligned, then a centered
  // title. The running "LastName PageNumber" head comes from the reference doc.
  generateHeaderFilter(yaml) {
    return `${this.luaBibHelpers()}
${this.luaMetaVars(yaml)}

function Pandoc(doc)
  local new_blocks = {}

  local info = { M.author, M.instructor, M.course,
                 M.date ~= "" and format_date(M.date, true) or "" }
  for _, line in ipairs(info) do
    if line ~= "" then
      table.insert(new_blocks, para_xml(line, 0, false))
    end
  end

  if M.title ~= "" then
    -- MLA keeps title and subtitle on one line, joined by a colon.
    local full_title = M.title
    if M.subtitle ~= "" then
      if M.title:match(":%s*$") then
        full_title = M.title:gsub("%s*$", "") .. " " .. M.subtitle
      else
        full_title = M.title .. ": " .. M.subtitle
      end
    end
    table.insert(new_blocks, para_xml(full_title, 0, true))
  end

  for _, block in ipairs(doc.blocks) do
    if not handle_bib_block(block, new_blocks) then
      table.insert(new_blocks, block)
    end
  end

  doc.blocks = new_blocks
  return doc
end`;
  }
  runCommand(stage, command, args, cwd) {
    return new Promise((resolve, reject) => {
      const isWindows = os.platform() === "win32";
      let proc;
      if (isWindows) {
        const quotedCommand = `"${command}"`;
        const quotedArgs = args.map((arg) => arg.includes(" ") ? `"${arg}"` : arg);
        proc = (0, import_child_process.spawn)(quotedCommand, quotedArgs, { cwd, shell: true });
      } else {
        proc = (0, import_child_process.spawn)(command, args, { cwd });
      }
      let stderr = "";
      let stdout = "";
      const timer = setTimeout(() => proc.kill(), COMMAND_TIMEOUT_MS);
      proc.stdout.on("data", (data) => {
        stdout += data.toString();
      });
      proc.stderr.on("data", (data) => {
        stderr += data.toString();
      });
      proc.on("close", (code, signal) => {
        clearTimeout(timer);
        if (code === 0) {
          resolve();
        } else {
          if (signal) stderr += `
killed by ${signal}`;
          reject(new CommandError(stage, [command, ...args], code, stdout, stderr));
        }
      });
      proc.on("error", (err) => {
        clearTimeout(timer);
        reject(err);
      });
    });
  }
  onunload() {
  }
};
function expandHome(p) {
  return p.startsWith("~") ? path.join(os.homedir(), p.slice(1)) : p;
}
function isFile(p) {
  try {
    return fs.statSync(p).isFile();
  } catch (e) {
    return false;
  }
}
function findBibFile(vaultPath) {
  const sources = path.join(vaultPath, "sources");
  try {
    const hit = fs.readdirSync(sources).sort().find((f) => f.endsWith(".bib") && !f.startsWith("._"));
    if (hit) return path.join(sources, hit);
  } catch (e) {
  }
  const walk = (dir) => {
    let entries;
    try {
      entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch (e) {
      return null;
    }
    entries.sort((a, b) => a.name.localeCompare(b.name));
    for (const e of entries) {
      if (e.isFile() && e.name.endsWith(".bib") && !e.name.startsWith("._")) return path.join(dir, e.name);
    }
    for (const e of entries) {
      if (e.isDirectory() && !e.name.startsWith(".") && e.name !== "node_modules") {
        const found = walk(path.join(dir, e.name));
        if (found) return found;
      }
    }
    return null;
  };
  return walk(vaultPath);
}
function luaStr(s) {
  return '"' + s.replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r?\n/g, "\\n") + '"';
}
function typstStr(s) {
  return '"' + s.replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"';
}
function authorLastname(yaml) {
  if (yaml["lastname"]) return yaml["lastname"];
  const parts = (yaml["author"] || "").split(/\s+/).filter(Boolean);
  return parts.length ? parts[parts.length - 1] : "";
}
function formatExportDate(yaml) {
  const raw = yaml["date"] || "";
  const m = raw.match(/(\d+)-(\d+)-(\d+)/);
  if (!m) return raw;
  const month = parseInt(m[2], 10);
  const day = parseInt(m[3], 10);
  if (month < 1 || month > 12) return raw;
  const name = MONTH_NAMES[month - 1];
  return styleKey(yaml) === "mla" ? `${day} ${name} ${m[1]}` : `${name} ${day}, ${m[1]}`;
}
function stripFrontmatter(content) {
  const m = content.match(/^---\r?\n[\s\S]*?\r?\n---[ \t]*\r?\n?/);
  return m ? content.slice(m[0].length) : content;
}
function typstWrapper(yaml, bodyName) {
  const fields = [
    ["title", yaml["title"] || ""],
    ["subtitle", yaml["subtitle"] || ""],
    ["author", yaml["author"] || ""],
    ["course", yaml["course"] || ""],
    ["instructor", yaml["instructor"] || ""],
    ["date", formatExportDate(yaml)],
    ["style", yaml["style"] || ""],
    ["spacing", yaml["spacing"] || "double"],
    ["lastname", authorLastname(yaml)]
  ];
  const args = fields.map(([k, v]) => `  ${k}: ${typstStr(v)},
`).join("");
  return `#import "folio.typ": conf
#show: conf.with(
${args})
#include ${typstStr(bodyName)}
`;
}
var BIB_HEADING_RE = /^(#{1,6})\s*(bibliography|works\s+cited|references)\s*$/im;
function wrapBibliography(markdown) {
  const m = BIB_HEADING_RE.exec(markdown);
  if (!m) return markdown;
  const start = m.index + m[0].length;
  const rest = markdown.slice(start);
  const nxt = /^#{1,6}\s/m.exec(rest);
  const end = start + (nxt ? nxt.index : rest.length);
  const body = markdown.slice(start, end);
  if (!body.trim() || body.includes(":::")) return markdown;
  return markdown.slice(0, start) + '\n\n::: {#bibentries custom-style="Bibliography"}\n' + body.replace(/^\n+|\n+$/g, "") + "\n:::\n" + markdown.slice(end);
}
var CITEKEY_RE = /@([A-Za-z0-9_][A-Za-z0-9_:.#$%&+?<>~/-]*)/g;
var BIB_GLOBALS = ["string", "preamble"];
function citedKeys(markdown) {
  const keys = /* @__PURE__ */ new Set();
  for (const m of markdown.matchAll(CITEKEY_RE)) keys.add(m[1].replace(/[.,;:]+$/, ""));
  return keys;
}
function bibBlocks(text) {
  const blocks = [];
  let i = 0;
  const n = text.length;
  while (true) {
    const at = text.indexOf("@", i);
    if (at < 0) break;
    const m = /^@\s*([A-Za-z]+)\s*([{(])/.exec(text.slice(at, at + 64));
    if (!m) {
      i = at + 1;
      continue;
    }
    const kind = m[1].toLowerCase();
    const opener = m[2];
    const closer = opener === "{" ? "}" : ")";
    let depth = 0;
    let j = at + m[0].length - 1;
    while (j < n) {
      const c = text[j];
      if (c === "\\") {
        j += 2;
        continue;
      }
      if (c === opener) depth++;
      else if (c === closer) {
        depth--;
        if (depth === 0) {
          j++;
          break;
        }
      }
      j++;
    }
    const body = text.slice(at, j);
    let key = null;
    if (!BIB_GLOBALS.includes(kind)) {
      const km = /^@\s*[A-Za-z]+\s*[{(]\s*([^,\s}]+)/.exec(body);
      key = km ? km[1] : null;
    }
    blocks.push({ kind, key, body });
    i = j;
  }
  return blocks;
}
function filterBib(text, keys) {
  var _a;
  const blocks = bibBlocks(text);
  if (!blocks.length) return null;
  const byKey = /* @__PURE__ */ new Map();
  for (const b of blocks) {
    if (!b.key) continue;
    if (!byKey.has(b.key)) byKey.set(b.key, b.body);
    if (!byKey.has(b.key.toLowerCase())) byKey.set(b.key.toLowerCase(), b.body);
  }
  const wanted = /* @__PURE__ */ new Set();
  const queue = [...keys];
  while (queue.length) {
    const k = queue.pop();
    const body = (_a = byKey.get(k)) != null ? _a : byKey.get(k.toLowerCase());
    if (body === void 0 || wanted.has(k)) continue;
    wanted.add(k);
    for (const m of body.matchAll(/(?:crossref|xdata)\s*=\s*[{"]([^}"]+)/gi)) {
      queue.push(...m[1].split(",").map((r) => r.trim()));
    }
  }
  if (!wanted.size) return null;
  const wantedLower = new Set([...wanted].map((w) => w.toLowerCase()));
  const out = blocks.filter((b) => BIB_GLOBALS.includes(b.kind) || b.key && wantedLower.has(b.key.toLowerCase())).map((b) => b.body);
  const result = out.join("\n\n") + "\n";
  const kept = new Set(bibBlocks(result).filter((b) => b.key).map((b) => b.key.toLowerCase()));
  for (const k of keys) {
    const present = byKey.has(k) || byKey.has(k.toLowerCase());
    if (present && !kept.has(k.toLowerCase())) return null;
  }
  return result;
}
var BIB_NOISE_RE = /([,{]\s*)(?:abstract|file|keywords)\s*=\s*/gi;
function bibValueEnd(text, i) {
  const n = text.length;
  if (i >= n) return null;
  if (text[i] === "{") {
    let depth = 0;
    while (i < n) {
      const c = text[i];
      if (c === "\\") {
        i += 2;
        continue;
      }
      if (c === "{") depth++;
      else if (c === "}") {
        depth--;
        if (depth === 0) return i + 1;
      }
      i++;
    }
    return null;
  }
  if (text[i] === '"') {
    i++;
    while (i < n) {
      if (text[i] === "\\") {
        i += 2;
        continue;
      }
      if (text[i] === '"') return i + 1;
      i++;
    }
    return null;
  }
  while (i < n && text[i] !== "," && text[i] !== "}") i++;
  return i;
}
function stripBibNoise(text) {
  const out = [];
  let pos = 0;
  let removed = 0;
  while (true) {
    BIB_NOISE_RE.lastIndex = pos;
    const m = BIB_NOISE_RE.exec(text);
    if (!m) {
      out.push(text.slice(pos));
      break;
    }
    const valueStart = m.index + m[0].length;
    let end = bibValueEnd(text, valueStart);
    if (end === null) {
      out.push(text.slice(pos, valueStart));
      pos = valueStart;
      continue;
    }
    out.push(text.slice(pos, m.index));
    if (m[1][0] === "{") {
      out.push("{");
      while (end < text.length && (text[end] === " " || text[end] === "	")) end++;
      if (end < text.length && text[end] === ",") end++;
    }
    pos = end;
    removed++;
  }
  return [out.join(""), removed];
}
function bibOpeners(text) {
  return [...text.matchAll(/@\s*[A-Za-z]+\s*[{(]\s*([^,\s}]+)/g)].map((m) => m[1]);
}
function narrowBib(bibSrc, markdown, tmpDir) {
  let text;
  try {
    text = fs.readFileSync(bibSrc, "utf8");
  } catch (e) {
    return bibSrc;
  }
  const keys = citedKeys(markdown);
  if (!keys.size) return bibSrc;
  let small = filterBib(text, keys);
  if (small) {
    small = stripBibNoise(small)[0];
  } else {
    const [stripped, removed] = stripBibNoise(text);
    const before = bibOpeners(text);
    const after = bibOpeners(stripped);
    if (!removed || before.length !== after.length || before.some((k, i) => k !== after[i])) return bibSrc;
    small = stripped;
  }
  if (small.length >= text.length * 0.9) return bibSrc;
  try {
    const dest = path.join(tmpDir, "cited.bib");
    fs.writeFileSync(dest, small, "utf8");
    return dest;
  } catch (e) {
    return bibSrc;
  }
}
var EMPTY_HEADER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Header"/>
    </w:pPr>
  </w:p>
</w:hdr>`;
var EMPTY_FOOTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Footer"/>
    </w:pPr>
  </w:p>
</w:ftr>`;
var CHAR_TWIPS = 125;
var CELL_PAD_TWIPS = 216;
var TEXT_WIDTH_TWIPS = 9360;
function columnWidths(rows, total = TEXT_WIDTH_TWIPS) {
  const n = rows.reduce((mx, r) => Math.max(mx, r.length), 0);
  if (!n) return [];
  const wantC = new Array(n).fill(0);
  const needC = new Array(n).fill(0);
  for (const row of rows) {
    row.slice(0, n).forEach((cell, i) => {
      const longestWord = cell.split(/\s+/).reduce((mx, w) => Math.max(mx, w.length), 1);
      needC[i] = Math.max(needC[i], longestWord);
      wantC[i] = Math.max(wantC[i], cell.length);
    });
  }
  const toTw = (c) => c * CHAR_TWIPS + CELL_PAD_TWIPS;
  const need = needC.map((c) => toTw(Math.max(3, c)));
  const want = need.map((nd, i) => Math.max(nd, toTw(wantC[i])));
  const out = new Array(n).fill(0);
  let remaining = total;
  const left = [...Array(n).keys()].sort((a, b) => want[a] - want[b]);
  while (left.length) {
    const share = Math.floor(remaining / left.length);
    const i = left[0];
    if (want[i] > share) {
      const tw = left.reduce((s, j) => s + want[j], 0) || 1;
      const alloc = new Map(left.map((j) => [j, Math.max(need[j], Math.floor(remaining * want[j] / tw))]));
      const over = [...alloc.values()].reduce((s, v) => s + v, 0);
      for (const [j, v] of alloc) {
        out[j] = over > remaining ? Math.max(1, Math.floor(v * remaining / over)) : v;
      }
      remaining = 0;
      break;
    }
    out[i] = want[i];
    remaining -= want[i];
    left.shift();
  }
  if (remaining > 0) out[out.indexOf(Math.max(...out))] += remaining;
  return out;
}
function cellTextRows(tableXml) {
  return [...tableXml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map((tr) => [...tr[0].matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map((tc) => [...tc[0].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((t) => t[1]).join("")));
}
function autofitTables(documentXml) {
  return documentXml.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, (tbl) => {
    const widths = columnWidths(cellTextRows(tbl));
    if (!widths.length) return tbl;
    let gi = 0;
    tbl = tbl.replace(
      /<w:gridCol w:w="\d+"\s*\/>/g,
      () => `<w:gridCol w:w="${widths[Math.min(gi++, widths.length - 1)]}"/>`
    );
    let ci = 0;
    tbl = tbl.replace(
      /<w:tcW w:type="dxa" w:w="\d+"\s*\/>/g,
      () => `<w:tcW w:type="dxa" w:w="${widths[ci++ % widths.length]}"/>`
    );
    return tbl;
  });
}
var AfterExportModal = class extends import_obsidian.Modal {
  constructor(app, outPath) {
    super(app);
    this.outPath = outPath;
  }
  onOpen() {
    const { contentEl } = this;
    const { shell } = require("electron");
    contentEl.createEl("h3", { text: "Export complete" });
    contentEl.createEl("p", { text: this.outPath });
    new import_obsidian.Setting(contentEl).addButton((btn) => btn.setButtonText("Open").setCta().onClick(async () => {
      const err = await shell.openPath(this.outPath);
      if (err) new import_obsidian.Notice(`Could not open ${path.basename(this.outPath)}: ${err}`);
      this.close();
    })).addButton((btn) => btn.setButtonText("Show in folder").onClick(() => {
      shell.showItemInFolder(this.outPath);
      this.close();
    })).addButton((btn) => btn.setButtonText("Close").onClick(() => this.close()));
  }
  onClose() {
    this.contentEl.empty();
  }
};
var MD2PDFSettingTab = class extends import_obsidian.PluginSettingTab {
  constructor(app, plugin) {
    super(app, plugin);
    this.plugin = plugin;
  }
  display() {
    const { containerEl } = this;
    containerEl.empty();
    containerEl.createEl("h2", { text: "MD to PDF Settings" });
    new import_obsidian.Setting(containerEl).setName("PDF Engine").setDesc("Typst renders from the bundled MLA/Turabian templates and needs no LibreOffice. Notes whose spacing: names a reference doc other than single/double always use LibreOffice, as does everything when Typst is not installed.").addDropdown((dropdown) => dropdown.addOption("typst", "Typst").addOption("libreoffice", "LibreOffice").setValue(this.plugin.settings.pdfEngine).onChange(async (value) => {
      this.plugin.settings.pdfEngine = value;
      await this.plugin.saveSettings();
    }));
    const availableDocs = this.plugin.getAvailableRefDocs();
    new import_obsidian.Setting(containerEl).setName("Default Reference Document").setDesc("Default typography/spacing template when no spacing is specified in YAML (DOCX and LibreOffice exports)").addDropdown((dropdown) => {
      dropdown.addOption("", "(none - use double, or first available)");
      for (const doc of availableDocs) {
        dropdown.addOption(doc, doc);
      }
      dropdown.setValue(this.plugin.settings.defaultRef);
      dropdown.onChange(async (value) => {
        this.plugin.settings.defaultRef = value;
        await this.plugin.saveSettings();
      });
    });
    new import_obsidian.Setting(containerEl).setName("Refresh Reference Documents").setDesc(availableDocs.length > 0 ? `Rescan plugin folder for .docx files \u2014 scanning ${this.plugin.getPluginFolder()}` : `No .docx files found in ${this.plugin.getPluginFolder()}`).addButton((btn) => btn.setButtonText("Refresh").onClick(() => {
      this.display();
      new import_obsidian.Notice("Reference documents refreshed");
    }));
    const pathSetting = (name, key, placeholder) => new import_obsidian.Setting(containerEl).setName(name).setDesc(`Full path to the ${name.replace(" Path", "")} executable`).addText((text) => text.setPlaceholder(placeholder).setValue(this.plugin.settings[key]).onChange(async (value) => {
      this.plugin.settings[key] = value;
      await this.plugin.saveSettings();
    }));
    pathSetting("Pandoc Path", "pandocPath", "/usr/bin/pandoc");
    pathSetting("Typst Path", "typstPath", "/usr/bin/typst");
    pathSetting("LibreOffice Path", "libreOfficePath", "/usr/bin/libreoffice");
    new import_obsidian.Setting(containerEl).setName("Auto-detect Paths").setDesc("Attempt to automatically find Pandoc, Typst and LibreOffice").addButton((btn) => btn.setButtonText("Auto-detect").onClick(async () => {
      this.plugin.settings.pandocPath = "";
      this.plugin.settings.libreOfficePath = "";
      this.plugin.settings.typstPath = "";
      await this.plugin.autoDetectPaths();
      this.display();
      new import_obsidian.Notice("Auto-detection complete");
    }));
    new import_obsidian.Setting(containerEl).setName("Custom Pandoc Arguments").setDesc("Additional arguments to pass to pandoc (space-separated)").addText((text) => text.setPlaceholder("--toc --number-sections").setValue(this.plugin.settings.customPandocArgs).onChange(async (value) => {
      this.plugin.settings.customPandocArgs = value;
      await this.plugin.saveSettings();
    }));
    new import_obsidian.Setting(containerEl).setName("Auto-detect Bibliography").setDesc("Enable citeproc when a bibliography field is found in frontmatter. Without csl:, the style: picks Chicago notes-bibliography or MLA.").addToggle((toggle) => toggle.setValue(this.plugin.settings.autoDetectBibliography).onChange(async (value) => {
      this.plugin.settings.autoDetectBibliography = value;
      await this.plugin.saveSettings();
    }));
    new import_obsidian.Setting(containerEl).setName("Offer to Open Export").setDesc("After exporting, offer to open the file or show it in its folder").addToggle((toggle) => toggle.setValue(this.plugin.settings.offerToOpen).onChange(async (value) => {
      this.plugin.settings.offerToOpen = value;
      await this.plugin.saveSettings();
    }));
    new import_obsidian.Setting(containerEl).setName("Delete Intermediate .docx").setDesc("Delete the intermediate .docx file after a LibreOffice PDF conversion").addToggle((toggle) => toggle.setValue(this.plugin.settings.deleteIntermediateDocx).onChange(async (value) => {
      this.plugin.settings.deleteIntermediateDocx = value;
      await this.plugin.saveSettings();
    }));
    new import_obsidian.Setting(containerEl).setName("PDF Output Location").setDesc('Where to save PDF files. Leave empty for same folder as source. Can be relative to vault (e.g., "exports/pdf") or absolute path.').addText((text) => text.setPlaceholder("(same as source file)").setValue(this.plugin.settings.pdfOutputLocation).onChange(async (value) => {
      this.plugin.settings.pdfOutputLocation = value;
      await this.plugin.saveSettings();
    }));
    new import_obsidian.Setting(containerEl).setName("DOCX Output Location").setDesc("Where to save .docx exports (and intermediate .docx files when kept). Leave empty for same folder as source.").addText((text) => text.setPlaceholder("(same as source file)").setValue(this.plugin.settings.docxOutputLocation).onChange(async (value) => {
      this.plugin.settings.docxOutputLocation = value;
      await this.plugin.saveSettings();
    }));
    containerEl.createEl("h3", { text: "Style Types" });
    const formatsEl = containerEl.createEl("div");
    formatsEl.innerHTML = `
			<p>Use the <code>style:</code> YAML field to control document structure:</p>
			<ul>
				<li><code>style: chicago</code> (or <code>turabian</code>) \u2014 Turabian-style title page (title centered 1/3 down, author/course/date in bottom half)</li>
				<li><code>style: mla</code> \u2014 MLA-style header block (author, instructor, course, date left-aligned at top)</li>
				<li><em>(no style)</em> \u2014 Plain document with no special header or cover page</li>
			</ul>
			<p>Run <strong>Insert export frontmatter</strong> from the command palette to add the fields.</p>
		`;
    containerEl.createEl("h3", { text: "Reference Documents" });
    const refEl = containerEl.createEl("div");
    refEl.innerHTML = `
			<p>Use the <code>spacing:</code> YAML field to specify typography and spacing:</p>
			<ul>
				<li><code>single</code> and <code>double</code> work with both engines</li>
				<li>Other <code>.docx</code> files in the plugin folder are used by filename without extension and always render through LibreOffice</li>
			</ul>
			<p><strong>Available reference documents:</strong> ${availableDocs.length > 0 ? availableDocs.join(", ") : "(none found)"}</p>
		`;
    containerEl.createEl("h3", { text: "YAML Frontmatter Example" });
    const yamlEl = containerEl.createEl("div");
    yamlEl.innerHTML = `
			<pre>---
title: My Thesis Title
subtitle: An Optional Second Line
author: John Smith
course: English 626
instructor: Dr. Johnson
date: 2025-03-07
style: chicago
spacing: double
bibliography: sources/library.bib
---</pre>
		`;
  }
};
