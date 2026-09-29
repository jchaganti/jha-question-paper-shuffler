//
// The installer's password check.
//
// Windows Installer runs this as a custom action, twice: once from the password dialog, to
// light up the message under the box, and once from the install sequence itself, where a
// wrong password stops the installation. The second one is the one that matters - it runs
// whether the installer was double-clicked or started from a command line with /qn, so a
// silent install cannot slip past the dialog.
//
// What it compares is a SHA-256 digest, never the password: the salt and the digest come in
// as the AdminSalt and AdminDigest properties, which `scripts/build-msi.mjs` fills in from
// `src/main/AdminSecret.ts` so that the app and the installer can never disagree about what
// the password is. They are *private* properties - mixed case, not all capitals - which is
// what stops someone handing `msiexec` a digest of their own on the command line.
//
// Written in plain ES3 with no library calls, because this runs inside the Windows Script
// Host that Windows Installer hosts, not in Node. `tests/installerPassword.test.ts` runs the
// same SHA-256 here under cscript and checks it against Node's, digest for digest.
//

/* ---------------------------------------------------------------- SHA-256 -- */

function rightRotate(value, bits) {
  return ((value >>> bits) | (value << (32 - bits))) >>> 0;
}

var SHA256_K = [
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
  0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2
];

/** `bytes` is an array of numbers 0..255. Returns 64 lower-case hex characters. */
function sha256Hex(bytes) {
  var h = [
    0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19
  ];

  // Padding: the 0x80 terminator, zeroes, then the length in bits as a 64-bit big-endian.
  var message = [];
  var i;
  for (i = 0; i < bytes.length; i++) message[i] = bytes[i] & 0xff;
  var bitLength = bytes.length * 8;
  message[message.length] = 0x80;
  while (message.length % 64 !== 56) message[message.length] = 0;
  // A question paper's password is far short of 2^32 bits, so the high word is always zero.
  message[message.length] = 0;
  message[message.length] = 0;
  message[message.length] = 0;
  message[message.length] = 0;
  message[message.length] = (bitLength >>> 24) & 0xff;
  message[message.length] = (bitLength >>> 16) & 0xff;
  message[message.length] = (bitLength >>> 8) & 0xff;
  message[message.length] = bitLength & 0xff;

  var w = [];
  for (var block = 0; block < message.length; block += 64) {
    for (i = 0; i < 16; i++) {
      w[i] =
        ((message[block + i * 4] << 24) |
          (message[block + i * 4 + 1] << 16) |
          (message[block + i * 4 + 2] << 8) |
          message[block + i * 4 + 3]) >>>
        0;
    }
    for (i = 16; i < 64; i++) {
      var s0 = (rightRotate(w[i - 15], 7) ^ rightRotate(w[i - 15], 18) ^ (w[i - 15] >>> 3)) >>> 0;
      var s1 = (rightRotate(w[i - 2], 17) ^ rightRotate(w[i - 2], 19) ^ (w[i - 2] >>> 10)) >>> 0;
      w[i] = (((w[i - 16] + s0) >>> 0) + ((w[i - 7] + s1) >>> 0)) >>> 0;
    }

    var a = h[0], b = h[1], c = h[2], d = h[3], e = h[4], f = h[5], g = h[6], hh = h[7];
    for (i = 0; i < 64; i++) {
      var S1 = (rightRotate(e, 6) ^ rightRotate(e, 11) ^ rightRotate(e, 25)) >>> 0;
      var ch = ((e & f) ^ (~e & g)) >>> 0;
      var temp1 = (((((hh + S1) >>> 0) + ch) >>> 0) + ((SHA256_K[i] + w[i]) >>> 0)) >>> 0;
      var S0 = (rightRotate(a, 2) ^ rightRotate(a, 13) ^ rightRotate(a, 22)) >>> 0;
      var maj = ((a & b) ^ (a & c) ^ (b & c)) >>> 0;
      var temp2 = (S0 + maj) >>> 0;

      hh = g;
      g = f;
      f = e;
      e = (d + temp1) >>> 0;
      d = c;
      c = b;
      b = a;
      a = (temp1 + temp2) >>> 0;
    }

    h[0] = (h[0] + a) >>> 0;
    h[1] = (h[1] + b) >>> 0;
    h[2] = (h[2] + c) >>> 0;
    h[3] = (h[3] + d) >>> 0;
    h[4] = (h[4] + e) >>> 0;
    h[5] = (h[5] + f) >>> 0;
    h[6] = (h[6] + g) >>> 0;
    h[7] = (h[7] + hh) >>> 0;
  }

  var hex = '';
  for (i = 0; i < 8; i++) {
    var word = h[i];
    for (var shift = 28; shift >= 0; shift -= 4) hex += '0123456789abcdef'.charAt((word >>> shift) & 0xf);
  }
  return hex;
}

/* ----------------------------------------------------------- conversions -- */

/** The bytes of `text` as UTF-8, which is what the app hashes too. */
function utf8Bytes(text) {
  var bytes = [];
  for (var i = 0; i < text.length; i++) {
    var code = text.charCodeAt(i);
    // A character outside the basic plane arrives as a surrogate pair; put it back together
    // so that a password with an emoji in it hashes the same here as it does in the app.
    if (code >= 0xd800 && code <= 0xdbff && i + 1 < text.length) {
      var low = text.charCodeAt(i + 1);
      if (low >= 0xdc00 && low <= 0xdfff) {
        code = 0x10000 + ((code - 0xd800) << 10) + (low - 0xdc00);
        i++;
      }
    }
    if (code < 0x80) {
      bytes[bytes.length] = code;
    } else if (code < 0x800) {
      bytes[bytes.length] = 0xc0 | (code >> 6);
      bytes[bytes.length] = 0x80 | (code & 0x3f);
    } else if (code < 0x10000) {
      bytes[bytes.length] = 0xe0 | (code >> 12);
      bytes[bytes.length] = 0x80 | ((code >> 6) & 0x3f);
      bytes[bytes.length] = 0x80 | (code & 0x3f);
    } else {
      bytes[bytes.length] = 0xf0 | (code >> 18);
      bytes[bytes.length] = 0x80 | ((code >> 12) & 0x3f);
      bytes[bytes.length] = 0x80 | ((code >> 6) & 0x3f);
      bytes[bytes.length] = 0x80 | (code & 0x3f);
    }
  }
  return bytes;
}

/** The salt arrives as hex, because hex needs no decoder worth testing. */
function hexBytes(hex) {
  var bytes = [];
  for (var i = 0; i + 1 < hex.length; i += 2) bytes[bytes.length] = parseInt(hex.substr(i, 2), 16);
  return bytes;
}

/** The digest of a typed password, exactly as `adminDigestOf` computes it in the app. */
function adminDigestOf(saltHex, password) {
  var bytes = hexBytes(saltHex);
  var typed = utf8Bytes(password);
  for (var i = 0; i < typed.length; i++) bytes[bytes.length] = typed[i];
  return sha256Hex(bytes);
}

/* ------------------------------------------------------- custom actions -- */

var msiDoActionStatusSuccess = 1;
var msiDoActionStatusFailure = 3;
var msiMessageTypeError = 0x01000000;

function passwordIsRight() {
  var salt = Session.Property('AdminSalt');
  var expected = Session.Property('AdminDigest');
  var typed = Session.Property('ADMINPASSWORD');
  if (!salt || !expected) return false;
  return adminDigestOf(salt, typed ? typed : '') === expected.toLowerCase();
}

/**
 * The dialog's check. Never fails the installation - it only decides whether the Next
 * button moves on, and what the line under the password box says.
 */
function CheckAdminPassword() {
  if (passwordIsRight()) {
    Session.Property('ADMINPASSWORDOK') = '1';
    Session.Property('ADMINPASSWORDMESSAGE') = '';
  } else {
    Session.Property('ADMINPASSWORDOK') = '0';
    Session.Property('ADMINPASSWORDMESSAGE') =
      'That is not the administrator password. Ask whoever sent you this installer for it.';
  }
  return msiDoActionStatusSuccess;
}

/**
 * The real gate, in the install sequence rather than the dialog.
 *
 * This is what a command-line install runs into: `msiexec /i ... /qn` never opens a dialog,
 * so without this there would be nothing to get past. The password has to arrive as
 * ADMINPASSWORD either way.
 */
function RequireAdminPassword() {
  if (passwordIsRight()) return msiDoActionStatusSuccess;
  var record = Session.Installer.CreateRecord(0);
  record.StringData(0) =
    'Question Paper Shuffler can only be installed by an administrator. ' +
    'Run the installer and type the administrator password, or pass it as ADMINPASSWORD=...';
  Session.Message(msiMessageTypeError, record);
  return msiDoActionStatusFailure;
}
