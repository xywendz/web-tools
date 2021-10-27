(function(f){if(typeof exports==="object"&&typeof module!=="undefined"){module.exports=f()}else if(typeof define==="function"&&define.amd){define([],f)}else{var g;if(typeof window!=="undefined"){g=window}else if(typeof global!=="undefined"){g=global}else if(typeof self!=="undefined"){g=self}else{g=this}g.LuckyExcel = f()}})(function(){var define,module,exports;return (function(){function r(e,n,t){function o(i,f){if(!n[i]){if(!e[i]){var c="function"==typeof require&&require;if(!f&&c)return c(i,!0);if(u)return u(i,!0);var a=new Error("Cannot find module '"+i+"'");throw a.code="MODULE_NOT_FOUND",a}var p=n[i]={exports:{}};e[i][0].call(p.exports,function(r){var n=e[i][1][r];return o(n||r)},p,p.exports,r,e,n,t)}return n[i].exports}for(var u="function"==typeof require&&require,i=0;i<t.length;i++)o(t[i]);return o}return r})()({1:[function(require,module,exports){
'use strict'

exports.byteLength = byteLength
exports.toByteArray = toByteArray
exports.fromByteArray = fromByteArray

var lookup = []
var revLookup = []
var Arr = typeof Uint8Array !== 'undefined' ? Uint8Array : Array

var code = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'
for (var i = 0, len = code.length; i < len; ++i) {
  lookup[i] = code[i]
  revLookup[code.charCodeAt(i)] = i
}

// Support decoding URL-safe base64 strings, as Node.js does.
// See: https://en.wikipedia.org/wiki/Base64#URL_applications
revLookup['-'.charCodeAt(0)] = 62
revLookup['_'.charCodeAt(0)] = 63

function getLens (b64) {
  var len = b64.length

  if (len % 4 > 0) {
    throw new Error('Invalid string. Length must be a multiple of 4')
  }

  // Trim off extra bytes after placeholder bytes are found
  // See: https://github.com/beatgammit/base64-js/issues/42
  var validLen = b64.indexOf('=')
  if (validLen === -1) validLen = len

  var placeHoldersLen = validLen === len
    ? 0
    : 4 - (validLen % 4)

  return [validLen, placeHoldersLen]
}

// base64 is 4/3 + up to two characters of the original data
function byteLength (b64) {
  var lens = getLens(b64)
  var validLen = lens[0]
  var placeHoldersLen = lens[1]
  return ((validLen + placeHoldersLen) * 3 / 4) - placeHoldersLen
}

function _byteLength (b64, validLen, placeHoldersLen) {
  return ((validLen + placeHoldersLen) * 3 / 4) - placeHoldersLen
}

function toByteArray (b64) {
  var tmp
  var lens = getLens(b64)
  var validLen = lens[0]
  var placeHoldersLen = lens[1]

  var arr = new Arr(_byteLength(b64, validLen, placeHoldersLen))

  var curByte = 0

  // if there are placeholders, only get up to the last complete 4 chars
  var len = placeHoldersLen > 0
    ? validLen - 4
    : validLen

  var i
  for (i = 0; i < len; i += 4) {
    tmp =
      (revLookup[b64.charCodeAt(i)] << 18) |
      (revLookup[b64.charCodeAt(i + 1)] << 12) |
      (revLookup[b64.charCodeAt(i + 2)] << 6) |
      revLookup[b64.charCodeAt(i + 3)]
    arr[curByte++] = (tmp >> 16) & 0xFF
    arr[curByte++] = (tmp >> 8) & 0xFF
    arr[curByte++] = tmp & 0xFF
  }

  if (placeHoldersLen === 2) {
    tmp =
      (revLookup[b64.charCodeAt(i)] << 2) |
      (revLookup[b64.charCodeAt(i + 1)] >> 4)
    arr[curByte++] = tmp & 0xFF
  }

  if (placeHoldersLen === 1) {
    tmp =
      (revLookup[b64.charCodeAt(i)] << 10) |
      (revLookup[b64.charCodeAt(i + 1)] << 4) |
      (revLookup[b64.charCodeAt(i + 2)] >> 2)
    arr[curByte++] = (tmp >> 8) & 0xFF
    arr[curByte++] = tmp & 0xFF
  }

  return arr
}

function tripletToBase64 (num) {
  return lookup[num >> 18 & 0x3F] +
    lookup[num >> 12 & 0x3F] +
    lookup[num >> 6 & 0x3F] +
    lookup[num & 0x3F]
}

function encodeChunk (uint8, start, end) {
  var tmp
  var output = []
  for (var i = start; i < end; i += 3) {
    tmp =
      ((uint8[i] << 16) & 0xFF0000) +
      ((uint8[i + 1] << 8) & 0xFF00) +
      (uint8[i + 2] & 0xFF)
    output.push(tripletToBase64(tmp))
  }
  return output.join('')
}

function fromByteArray (uint8) {
  var tmp
  var len = uint8.length
  var extraBytes = len % 3 // if we have 1 byte left, pad 2 bytes
  var parts = []
  var maxChunkLength = 16383 // must be multiple of 3

  // go through the array every three bytes, we'll deal with trailing stuff later
  for (var i = 0, len2 = len - extraBytes; i < len2; i += maxChunkLength) {
    parts.push(encodeChunk(uint8, i, (i + maxChunkLength) > len2 ? len2 : (i + maxChunkLength)))
  }

  // pad the end with zeros, but make sure to not forget the extra bytes
  if (extraBytes === 1) {
    tmp = uint8[len - 1]
    parts.push(
      lookup[tmp >> 2] +
      lookup[(tmp << 4) & 0x3F] +
      '=='
    )
  } else if (extraBytes === 2) {
    tmp = (uint8[len - 2] << 8) + uint8[len - 1]
    parts.push(
      lookup[tmp >> 10] +
      lookup[(tmp >> 4) & 0x3F] +
      lookup[(tmp << 2) & 0x3F] +
      '='
    )
  }

  return parts.join('')
}

},{}],2:[function(require,module,exports){
(function (global,Buffer){(function (){
/*!
 * The buffer module from node.js, for the browser.
 *
 * @author   Feross Aboukhadijeh <http://feross.org>
 * @license  MIT
 */
/* eslint-disable no-proto */

'use strict'

var base64 = require('base64-js')
var ieee754 = require('ieee754')
var isArray = require('isarray')

exports.Buffer = Buffer
exports.SlowBuffer = SlowBuffer
exports.INSPECT_MAX_BYTES = 50

/**
 * If `Buffer.TYPED_ARRAY_SUPPORT`:
 *   === true    Use Uint8Array implementation (fastest)
 *   === false   Use Object implementation (most compatible, even IE6)
 *
 * Browsers that support typed arrays are IE 10+, Firefox 4+, Chrome 7+, Safari 5.1+,
 * Opera 11.6+, iOS 4.2+.
 *
 * Due to various browser bugs, sometimes the Object implementation will be used even
 * when the browser supports typed arrays.
 *
 * Note:
 *
 *   - Firefox 4-29 lacks support for adding new properties to `Uint8Array` instances,
 *     See: https://bugzilla.mozilla.org/show_bug.cgi?id=695438.
 *
 *   - Chrome 9-10 is missing the `TypedArray.prototype.subarray` function.
 *
 *   - IE10 has a broken `TypedArray.prototype.subarray` function which returns arrays of
 *     incorrect length in some situations.

 * We detect these buggy browsers and set `Buffer.TYPED_ARRAY_SUPPORT` to `false` so they
 * get the Object implementation, which is slower but behaves correctly.
 */
Buffer.TYPED_ARRAY_SUPPORT = global.TYPED_ARRAY_SUPPORT !== undefined
  ? global.TYPED_ARRAY_SUPPORT
  : typedArraySupport()

/*
 * Export kMaxLength after typed array support is determined.
 */
exports.kMaxLength = kMaxLength()

function typedArraySupport () {
  try {
    var arr = new Uint8Array(1)
    arr.__proto__ = {__proto__: Uint8Array.prototype, foo: function () { return 42 }}
    return arr.foo() === 42 && // typed array instances can be augmented
        typeof arr.subarray === 'function' && // chrome 9-10 lack `subarray`
        arr.subarray(1, 1).byteLength === 0 // ie10 has broken `subarray`
  } catch (e) {
    return false
  }
}

function kMaxLength () {
  return Buffer.TYPED_ARRAY_SUPPORT
    ? 0x7fffffff
    : 0x3fffffff
}

function createBuffer (that, length) {
  if (kMaxLength() < length) {
    throw new RangeError('Invalid typed array length')
  }
  if (Buffer.TYPED_ARRAY_SUPPORT) {
    // Return an augmented `Uint8Array` instance, for best performance
    that = new Uint8Array(length)
    that.__proto__ = Buffer.prototype
  } else {
    // Fallback: Return an object instance of the Buffer class
    if (that === null) {
      that = new Buffer(length)
    }
    that.length = length
  }

  return that
}

/**
 * The Buffer constructor returns instances of `Uint8Array` that have their
 * prototype changed to `Buffer.prototype`. Furthermore, `Buffer` is a subclass of
 * `Uint8Array`, so the returned instances will have all the node `Buffer` methods
 * and the `Uint8Array` methods. Square bracket notation works as expected -- it
 * returns a single octet.
 *
 * The `Uint8Array` prototype remains unmodified.
 */

function Buffer (arg, encodingOrOffset, length) {
  if (!Buffer.TYPED_ARRAY_SUPPORT && !(this instanceof Buffer)) {
    return new Buffer(arg, encodingOrOffset, length)
  }

  // Common case.
  if (typeof arg === 'number') {
    if (typeof encodingOrOffset === 'string') {
      throw new Error(
        'If encoding is specified then the first argument must be a string'
      )
    }
    return allocUnsafe(this, arg)
  }
  return from(this, arg, encodingOrOffset, length)
}

Buffer.poolSize = 8192 // not used by this implementation

// TODO: Legacy, not needed anymore. Remove in next major version.
Buffer._augment = function (arr) {
  arr.__proto__ = Buffer.prototype
  return arr
}

function from (that, value, encodingOrOffset, length) {
  if (typeof value === 'number') {
    throw new TypeError('"value" argument must not be a number')
  }

  if (typeof ArrayBuffer !== 'undefined' && value instanceof ArrayBuffer) {
    return fromArrayBuffer(that, value, encodingOrOffset, length)
  }

  if (typeof value === 'string') {
    return fromString(that, value, encodingOrOffset)
  }

  return fromObject(that, value)
}

/**
 * Functionally equivalent to Buffer(arg, encoding) but throws a TypeError
 * if value is a number.
 * Buffer.from(str[, encoding])
 * Buffer.from(array)
 * Buffer.from(buffer)
 * Buffer.from(arrayBuffer[, byteOffset[, length]])
 **/
Buffer.from = function (value, encodingOrOffset, length) {
  return from(null, value, encodingOrOffset, length)
}

if (Buffer.TYPED_ARRAY_SUPPORT) {
  Buffer.prototype.__proto__ = Uint8Array.prototype
  Buffer.__proto__ = Uint8Array
  if (typeof Symbol !== 'undefined' && Symbol.species &&
      Buffer[Symbol.species] === Buffer) {
    // Fix subarray() in ES2016. See: https://github.com/feross/buffer/pull/97
    Object.defineProperty(Buffer, Symbol.species, {
      value: null,
      configurable: true
    })
  }
}

function assertSize (size) {
  if (typeof size !== 'number') {
    throw new TypeError('"size" argument must be a number')
  } else if (size < 0) {
    throw new RangeError('"size" argument must not be negative')
  }
}

function alloc (that, size, fill, encoding) {
  assertSize(size)
  if (size <= 0) {
    return createBuffer(that, size)
  }
  if (fill !== undefined) {
    // Only pay attention to encoding if it's a string. This
    // prevents accidentally sending in a number that would
    // be interpretted as a start offset.
    return typeof encoding === 'string'
      ? createBuffer(that, size).fill(fill, encoding)
      : createBuffer(that, size).fill(fill)
  }
  return createBuffer(that, size)
}

/**
 * Creates a new filled Buffer instance.
 * alloc(size[, fill[, encoding]])
 **/
Buffer.alloc = function (size, fill, encoding) {
  return alloc(null, size, fill, encoding)
}

function allocUnsafe (that, size) {
  assertSize(size)
  that = createBuffer(that, size < 0 ? 0 : checked(size) | 0)
  if (!Buffer.TYPED_ARRAY_SUPPORT) {
    for (var i = 0; i < size; ++i) {
      that[i] = 0
    }
  }
  return that
}

/**
 * Equivalent to Buffer(num), by default creates a non-zero-filled Buffer instance.
 * */
Buffer.allocUnsafe = function (size) {
  return allocUnsafe(null, size)
}
/**
 * Equivalent to SlowBuffer(num), by default creates a non-zero-filled Buffer instance.
 */
Buffer.allocUnsafeSlow = function (size) {
  return allocUnsafe(null, size)
}

function fromString (that, string, encoding) {
  if (typeof encoding !== 'string' || encoding === '') {
    encoding = 'utf8'
  }

  if (!Buffer.isEncoding(encoding)) {
    throw new TypeError('"encoding" must be a valid string encoding')
  }

  var length = byteLength(string, encoding) | 0
  that = createBuffer(that, length)

  var actual = that.write(string, encoding)

  if (actual !== length) {
    // Writing a hex string, for example, that contains invalid characters will
    // cause everything after the first invalid character to be ignored. (e.g.
    // 'abxxcd' will be treated as 'ab')
    that = that.slice(0, actual)
  }

  return that
}

function fromArrayLike (that, array) {
  var length = array.length < 0 ? 0 : checked(array.length) | 0
  that = createBuffer(that, length)
  for (var i = 0; i < length; i += 1) {
    that[i] = array[i] & 255
  }
  return that
}

function fromArrayBuffer (that, array, byteOffset, length) {
  array.byteLength // this throws if `array` is not a valid ArrayBuffer

  if (byteOffset < 0 || array.byteLength < byteOffset) {
    throw new RangeError('\'offset\' is out of bounds')
  }

  if (array.byteLength < byteOffset + (length || 0)) {
    throw new RangeError('\'length\' is out of bounds')
  }

  if (byteOffset === undefined && length === undefined) {
    array = new Uint8Array(array)
  } else if (length === undefined) {
    array = new Uint8Array(array, byteOffset)
  } else {
    array = new Uint8Array(array, byteOffset, length)
  }

  if (Buffer.TYPED_ARRAY_SUPPORT) {
    // Return an augmented `Uint8Array` instance, for best performance
    that = array
    that.__proto__ = Buffer.prototype
  } else {
    // Fallback: Return an object instance of the Buffer class
    that = fromArrayLike(that, array)
  }
  return that
}

function fromObject (that, obj) {
  if (Buffer.isBuffer(obj)) {
    var len = checked(obj.length) | 0
    that = createBuffer(that, len)

    if (that.length === 0) {
      return that
    }

    obj.copy(that, 0, 0, len)
    return that
  }

  if (obj) {
    if ((typeof ArrayBuffer !== 'undefined' &&
        obj.buffer instanceof ArrayBuffer) || 'length' in obj) {
      if (typeof obj.length !== 'number' || isnan(obj.length)) {
        return createBuffer(that, 0)
      }
      return fromArrayLike(that, obj)
    }

    if (obj.type === 'Buffer' && isArray(obj.data)) {
      return fromArrayLike(that, obj.data)
    }
  }

  throw new TypeError('First argument must be a string, Buffer, ArrayBuffer, Array, or array-like object.')
}

function checked (length) {
  // Note: cannot use `length < kMaxLength()` here because that fails when
  // length is NaN (which is otherwise coerced to zero.)
  if (length >= kMaxLength()) {
    throw new RangeError('Attempt to allocate Buffer larger than maximum ' +
                         'size: 0x' + kMaxLength().toString(16) + ' bytes')
  }
  return length | 0
}

function SlowBuffer (length) {
  if (+length != length) { // eslint-disable-line eqeqeq
    length = 0
  }
  return Buffer.alloc(+length)
}

Buffer.isBuffer = function isBuffer (b) {
  return !!(b != null && b._isBuffer)
}

Buffer.compare = function compare (a, b) {
  if (!Buffer.isBuffer(a) || !Buffer.isBuffer(b)) {
    throw new TypeError('Arguments must be Buffers')
  }

  if (a === b) return 0

  var x = a.length
  var y = b.length

  for (var i = 0, len = Math.min(x, y); i < len; ++i) {
    if (a[i] !== b[i]) {
      x = a[i]
      y = b[i]
      break
    }
  }

  if (x < y) return -1
  if (y < x) return 1
  return 0
}

Buffer.isEncoding = function isEncoding (encoding) {
  switch (String(encoding).toLowerCase()) {
    case 'hex':
    case 'utf8':
    case 'utf-8':
    case 'ascii':
    case 'latin1':
    case 'binary':
    case 'base64':
    case 'ucs2':
    case 'ucs-2':
    case 'utf16le':
    case 'utf-16le':
      return true
    default:
      return false
  }
}

Buffer.concat = function concat (list, length) {
  if (!isArray(list)) {
    throw new TypeError('"list" argument must be an Array of Buffers')
  }

  if (list.length === 0) {
    return Buffer.alloc(0)
  }

  var i
  if (length === undefined) {
    length = 0
    for (i = 0; i < list.length; ++i) {
      length += list[i].length
    }
  }

  var buffer = Buffer.allocUnsafe(length)
  var pos = 0
  for (i = 0; i < list.length; ++i) {
    var buf = list[i]
    if (!Buffer.isBuffer(buf)) {
      throw new TypeError('"list" argument must be an Array of Buffers')
    }
    buf.copy(buffer, pos)
    pos += buf.length
  }
  return buffer
}

function byteLength (string, encoding) {
  if (Buffer.isBuffer(string)) {
    return string.length
  }
  if (typeof ArrayBuffer !== 'undefined' && typeof ArrayBuffer.isView === 'function' &&
      (ArrayBuffer.isView(string) || string instanceof ArrayBuffer)) {
    return string.byteLength
  }
  if (typeof string !== 'string') {
    string = '' + string
  }

  var len = string.length
  if (len === 0) return 0

  // Use a for loop to avoid recursion
  var loweredCase = false
  for (;;) {
    switch (encoding) {
      case 'ascii':
      case 'latin1':
      case 'binary':
        return len
      case 'utf8':
      case 'utf-8':
      case undefined:
        return utf8ToBytes(string).length
      case 'ucs2':
      case 'ucs-2':
      case 'utf16le':
      case 'utf-16le':
        return len * 2
      case 'hex':
        return len >>> 1
      case 'base64':
        return base64ToBytes(string).length
      default:
        if (loweredCase) return utf8ToBytes(string).length // assume utf8
        encoding = ('' + encoding).toLowerCase()
        loweredCase = true
    }
  }
}
Buffer.byteLength = byteLength

function slowToString (encoding, start, end) {
  var loweredCase = false

  // No need to verify that "this.length <= MAX_UINT32" since it's a read-only
  // property of a typed array.

  // This behaves neither like String nor Uint8Array in that we set start/end
  // to their upper/lower bounds if the value passed is out of range.
  // undefined is handled specially as per ECMA-262 6th Edition,
  // Section 13.3.3.7 Runtime Semantics: KeyedBindingInitialization.
  if (start === undefined || start < 0) {
    start = 0
  }
  // Return early if start > this.length. Done here to prevent potential uint32
  // coercion fail below.
  if (start > this.length) {
    return ''
  }

  if (end === undefined || end > this.length) {
    end = this.length
  }

  if (end <= 0) {
    return ''
  }

  // Force coersion to uint32. This will also coerce falsey/NaN values to 0.
  end >>>= 0
  start >>>= 0

  if (end <= start) {
    return ''
  }

  if (!encoding) encoding = 'utf8'

  while (true) {
    switch (encoding) {
      case 'hex':
        return hexSlice(this, start, end)

      case 'utf8':
      case 'utf-8':
        return utf8Slice(this, start, end)

      case 'ascii':
        return asciiSlice(this, start, end)

      case 'latin1':
      case 'binary':
        return latin1Slice(this, start, end)

      case 'base64':
        return base64Slice(this, start, end)

      case 'ucs2':
      case 'ucs-2':
      case 'utf16le':
      case 'utf-16le':
        return utf16leSlice(this, start, end)

      default:
        if (loweredCase) throw new TypeError('Unknown encoding: ' + encoding)
        encoding = (encoding + '').toLowerCase()
        loweredCase = true
    }
  }
}

// The property is used by `Buffer.isBuffer` and `is-buffer` (in Safari 5-7) to detect
// Buffer instances.
Buffer.prototype._isBuffer = true

function swap (b, n, m) {
  var i = b[n]
  b[n] = b[m]
  b[m] = i
}

Buffer.prototype.swap16 = function swap16 () {
  var len = this.length
  if (len % 2 !== 0) {
    throw new RangeError('Buffer size must be a multiple of 16-bits')
  }
  for (var i = 0; i < len; i += 2) {
    swap(this, i, i + 1)
  }
  return this
}

Buffer.prototype.swap32 = function swap32 () {
  var len = this.length
  if (len % 4 !== 0) {
    throw new RangeError('Buffer size must be a multiple of 32-bits')
  }
  for (var i = 0; i < len; i += 4) {
    swap(this, i, i + 3)
    swap(this, i + 1, i + 2)
  }
  return this
}

Buffer.prototype.swap64 = function swap64 () {
  var len = this.length
  if (len % 8 !== 0) {
    throw new RangeError('Buffer size must be a multiple of 64-bits')
  }
  for (var i = 0; i < len; i += 8) {
    swap(this, i, i + 7)
    swap(this, i + 1, i + 6)
    swap(this, i + 2, i + 5)
    swap(this, i + 3, i + 4)
  }
  return this
}

Buffer.prototype.toString = function toString () {
  var length = this.length | 0
  if (length === 0) return ''
  if (arguments.length === 0) return utf8Slice(this, 0, length)
  return slowToString.apply(this, arguments)
}

Buffer.prototype.equals = function equals (b) {
  if (!Buffer.isBuffer(b)) throw new TypeError('Argument must be a Buffer')
  if (this === b) return true
  return Buffer.compare(this, b) === 0
}

Buffer.prototype.inspect = function inspect () {
  var str = ''
  var max = exports.INSPECT_MAX_BYTES
  if (this.length > 0) {
    str = this.toString('hex', 0, max).match(/.{2}/g).join(' ')
    if (this.length > max) str += ' ... '
  }
  return '<Buffer ' + str + '>'
}

Buffer.prototype.compare = function compare (target, start, end, thisStart, thisEnd) {
  if (!Buffer.isBuffer(target)) {
    throw new TypeError('Argument must be a Buffer')
  }

  if (start === undefined) {
    start = 0
  }
  if (end === undefined) {
    end = target ? target.length : 0
  }
  if (thisStart === undefined) {
    thisStart = 0
  }
  if (thisEnd === undefined) {
    thisEnd = this.length
  }

  if (start < 0 || end > target.length || thisStart < 0 || thisEnd > this.length) {
    throw new RangeError('out of range index')
  }

  if (thisStart >= thisEnd && start >= end) {
    return 0
  }
  if (thisStart >= thisEnd) {
    return -1
  }
  if (start >= end) {
    return 1
  }

  start >>>= 0
  end >>>= 0
  thisStart >>>= 0
  thisEnd >>>= 0

  if (this === target) return 0

  var x = thisEnd - thisStart
  var y = end - start
  var len = Math.min(x, y)

  var thisCopy = this.slice(thisStart, thisEnd)
  var targetCopy = target.slice(start, end)

  for (var i = 0; i < len; ++i) {
    if (thisCopy[i] !== targetCopy[i]) {
      x = thisCopy[i]
      y = targetCopy[i]
      break
    }
  }

  if (x < y) return -1
  if (y < x) return 1
  return 0
}

// Finds either the first index of `val` in `buffer` at offset >= `byteOffset`,
// OR the last index of `val` in `buffer` at offset <= `byteOffset`.
//
// Arguments:
// - buffer - a Buffer to search
// - val - a string, Buffer, or number
// - byteOffset - an index into `buffer`; will be clamped to an int32
// - encoding - an optional encoding, relevant is val is a string
// - dir - true for indexOf, false for lastIndexOf
function bidirectionalIndexOf (buffer, val, byteOffset, encoding, dir) {
  // Empty buffer means no match
  if (buffer.length === 0) return -1

  // Normalize byteOffset
  if (typeof byteOffset === 'string') {
    encoding = byteOffset
    byteOffset = 0
  } else if (byteOffset > 0x7fffffff) {
    byteOffset = 0x7fffffff
  } else if (byteOffset < -0x80000000) {
    byteOffset = -0x80000000
  }
  byteOffset = +byteOffset  // Coerce to Number.
  if (isNaN(byteOffset)) {
    // byteOffset: it it's undefined, null, NaN, "foo", etc, search whole buffer
    byteOffset = dir ? 0 : (buffer.length - 1)
  }

  // Normalize byteOffset: negative offsets start from the end of the buffer
  if (byteOffset < 0) byteOffset = buffer.length + byteOffset
  if (byteOffset >= buffer.length) {
    if (dir) return -1
    else byteOffset = buffer.length - 1
  } else if (byteOffset < 0) {
    if (dir) byteOffset = 0
    else return -1
  }

  // Normalize val
  if (typeof val === 'string') {
    val = Buffer.from(val, encoding)
  }

  // Finally, search either indexOf (if dir is true) or lastIndexOf
  if (Buffer.isBuffer(val)) {
    // Special case: looking for empty string/buffer always fails
    if (val.length === 0) {
      return -1
    }
    return arrayIndexOf(buffer, val, byteOffset, encoding, dir)
  } else if (typeof val === 'number') {
    val = val & 0xFF // Search for a byte value [0-255]
    if (Buffer.TYPED_ARRAY_SUPPORT &&
        typeof Uint8Array.prototype.indexOf === 'function') {
      if (dir) {
        return Uint8Array.prototype.indexOf.call(buffer, val, byteOffset)
      } else {
        return Uint8Array.prototype.lastIndexOf.call(buffer, val, byteOffset)
      }
    }
    return arrayIndexOf(buffer, [ val ], byteOffset, encoding, dir)
  }

  throw new TypeError('val must be string, number or Buffer')
}

function arrayIndexOf (arr, val, byteOffset, encoding, dir) {
  var indexSize = 1
  var arrLength = arr.length
  var valLength = val.length

  if (encoding !== undefined) {
    encoding = String(encoding).toLowerCase()
    if (encoding === 'ucs2' || encoding === 'ucs-2' ||
        encoding === 'utf16le' || encoding === 'utf-16le') {
      if (arr.length < 2 || val.length < 2) {
        return -1
      }
      indexSize = 2
      arrLength /= 2
      valLength /= 2
      byteOffset /= 2
    }
  }

  function read (buf, i) {
    if (indexSize === 1) {
      return buf[i]
    } else {
      return buf.readUInt16BE(i * indexSize)
    }
  }

  var i
  if (dir) {
    var foundIndex = -1
    for (i = byteOffset; i < arrLength; i++) {
      if (read(arr, i) === read(val, foundIndex === -1 ? 0 : i - foundIndex)) {
        if (foundIndex === -1) foundIndex = i
        if (i - foundIndex + 1 === valLength) return foundIndex * indexSize
      } else {
        if (foundIndex !== -1) i -= i - foundIndex
        foundIndex = -1
      }
    }
  } else {
    if (byteOffset + valLength > arrLength) byteOffset = arrLength - valLength
    for (i = byteOffset; i >= 0; i--) {
      var found = true
      for (var j = 0; j < valLength; j++) {
        if (read(arr, i + j) !== read(val, j)) {
          found = false
          break
        }
      }
      if (found) return i
    }
  }

  return -1
}

Buffer.prototype.includes = function includes (val, byteOffset, encoding) {
  return this.indexOf(val, byteOffset, encoding) !== -1
}

Buffer.prototype.indexOf = function indexOf (val, byteOffset, encoding) {
  return bidirectionalIndexOf(this, val, byteOffset, encoding, true)
}

Buffer.prototype.lastIndexOf = function lastIndexOf (val, byteOffset, encoding) {
  return bidirectionalIndexOf(this, val, byteOffset, encoding, false)
}

function hexWrite (buf, string, offset, length) {
  offset = Number(offset) || 0
  var remaining = buf.length - offset
  if (!length) {
    length = remaining
  } else {
    length = Number(length)
    if (length > remaining) {
      length = remaining
    }
  }

  // must be an even number of digits
  var strLen = string.length
  if (strLen % 2 !== 0) throw new TypeError('Invalid hex string')

  if (length > strLen / 2) {
    length = strLen / 2
  }
  for (var i = 0; i < length; ++i) {
    var parsed = parseInt(string.substr(i * 2, 2), 16)
    if (isNaN(parsed)) return i
    buf[offset + i] = parsed
  }
  return i
}

function utf8Write (buf, string, offset, length) {
  return blitBuffer(utf8ToBytes(string, buf.length - offset), buf, offset, length)
}

function asciiWrite (buf, string, offset, length) {
  return blitBuffer(asciiToBytes(string), buf, offset, length)
}

function latin1Write (buf, string, offset, length) {
  return asciiWrite(buf, string, offset, length)
}

function base64Write (buf, string, offset, length) {
  return blitBuffer(base64ToBytes(string), buf, offset, length)
}

function ucs2Write (buf, string, offset, length) {
  return blitBuffer(utf16leToBytes(string, buf.length - offset), buf, offset, length)
}

Buffer.prototype.write = function write (string, offset, length, encoding) {
  // Buffer#write(string)
  if (offset === undefined) {
    encoding = 'utf8'
    length = this.length
    offset = 0
  // Buffer#write(string, encoding)
  } else if (length === undefined && typeof offset === 'string') {
    encoding = offset
    length = this.length
    offset = 0
  // Buffer#write(string, offset[, length][, encoding])
  } else if (isFinite(offset)) {
    offset = offset | 0
    if (isFinite(length)) {
      length = length | 0
      if (encoding === undefined) encoding = 'utf8'
    } else {
      encoding = length
      length = undefined
    }
  // legacy write(string, encoding, offset, length) - remove in v0.13
  } else {
    throw new Error(
      'Buffer.write(string, encoding, offset[, length]) is no longer supported'
    )
  }

  var remaining = this.length - offset
  if (length === undefined || length > remaining) length = remaining

  if ((string.length > 0 && (length < 0 || offset < 0)) || offset > this.length) {
    throw new RangeError('Attempt to write outside buffer bounds')
  }

  if (!encoding) encoding = 'utf8'

  var loweredCase = false
  for (;;) {
    switch (encoding) {
      case 'hex':
        return hexWrite(this, string, offset, length)

      case 'utf8':
      case 'utf-8':
        return utf8Write(this, string, offset, length)

      case 'ascii':
        return asciiWrite(this, string, offset, length)

      case 'latin1':
      case 'binary':
        return latin1Write(this, string, offset, length)

      case 'base64':
        // Warning: maxLength not taken into account in base64Write
        return base64Write(this, string, offset, length)

      case 'ucs2':
      case 'ucs-2':
      case 'utf16le':
      case 'utf-16le':
        return ucs2Write(this, string, offset, length)

      default:
        if (loweredCase) throw new TypeError('Unknown encoding: ' + encoding)
        encoding = ('' + encoding).toLowerCase()
        loweredCase = true
    }
  }
}

Buffer.prototype.toJSON = function toJSON () {
  return {
    type: 'Buffer',
    data: Array.prototype.slice.call(this._arr || this, 0)
  }
}

function base64Slice (buf, start, end) {
  if (start === 0 && end === buf.length) {
    return base64.fromByteArray(buf)
  } else {
    return base64.fromByteArray(buf.slice(start, end))
  }
}

function utf8Slice (buf, start, end) {
  end = Math.min(buf.length, end)
  var res = []

  var i = start
  while (i < end) {
    var firstByte = buf[i]
    var codePoint = null
    var bytesPerSequence = (firstByte > 0xEF) ? 4
      : (firstByte > 0xDF) ? 3
      : (firstByte > 0xBF) ? 2
      : 1

    if (i + bytesPerSequence <= end) {
      var secondByte, thirdByte, fourthByte, tempCodePoint

      switch (bytesPerSequence) {
        case 1:
          if (firstByte < 0x80) {
            codePoint = firstByte
          }
          break
        case 2:
          secondByte = buf[i + 1]
          if ((secondByte & 0xC0) === 0x80) {
            tempCodePoint = (firstByte & 0x1F) << 0x6 | (secondByte & 0x3F)
            if (tempCodePoint > 0x7F) {
              codePoint = tempCodePoint
            }
          }
          break
        case 3:
          secondByte = buf[i + 1]
          thirdByte = buf[i + 2]
          if ((secondByte & 0xC0) === 0x80 && (thirdByte & 0xC0) === 0x80) {
            tempCodePoint = (firstByte & 0xF) << 0xC | (secondByte & 0x3F) << 0x6 | (thirdByte & 0x3F)
            if (tempCodePoint > 0x7FF && (tempCodePoint < 0xD800 || tempCodePoint > 0xDFFF)) {
              codePoint = tempCodePoint
            }
          }
          break
        case 4:
          secondByte = buf[i + 1]
          thirdByte = buf[i + 2]
          fourthByte = buf[i + 3]
          if ((secondByte & 0xC0) === 0x80 && (thirdByte & 0xC0) === 0x80 && (fourthByte & 0xC0) === 0x80) {
            tempCodePoint = (firstByte & 0xF) << 0x12 | (secondByte & 0x3F) << 0xC | (thirdByte & 0x3F) << 0x6 | (fourthByte & 0x3F)
            if (tempCodePoint > 0xFFFF && tempCodePoint < 0x110000) {
              codePoint = tempCodePoint
            }
          }
      }
    }

    if (codePoint === null) {
      // we did not generate a valid codePoint so insert a
      // replacement char (U+FFFD) and advance only 1 byte
      codePoint = 0xFFFD
      bytesPerSequence = 1
    } else if (codePoint > 0xFFFF) {
      // encode to utf16 (surrogate pair dance)
      codePoint -= 0x10000
      res.push(codePoint >>> 10 & 0x3FF | 0xD800)
      codePoint = 0xDC00 | codePoint & 0x3FF
    }

    res.push(codePoint)
    i += bytesPerSequence
  }

  return decodeCodePointsArray(res)
}

// Based on http://stackoverflow.com/a/22747272/680742, the browser with
// the lowest limit is Chrome, with 0x10000 args.
// We go 1 magnitude less, for safety
var MAX_ARGUMENTS_LENGTH = 0x1000

function decodeCodePointsArray (codePoints) {
  var len = codePoints.length
  if (len <= MAX_ARGUMENTS_LENGTH) {
    return String.fromCharCode.apply(String, codePoints) // avoid extra slice()
  }

  // Decode in chunks to avoid "call stack size exceeded".
  var res = ''
  var i = 0
  while (i < len) {
    res += String.fromCharCode.apply(
      String,
      codePoints.slice(i, i += MAX_ARGUMENTS_LENGTH)
    )
  }
  return res
}

function asciiSlice (buf, start, end) {
  var ret = ''
  end = Math.min(buf.length, end)

  for (var i = start; i < end; ++i) {
    ret += String.fromCharCode(buf[i] & 0x7F)
  }
  return ret
}

function latin1Slice (buf, start, end) {
  var ret = ''
  end = Math.min(buf.length, end)

  for (var i = start; i < end; ++i) {
    ret += String.fromCharCode(buf[i])
  }
  return ret
}

function hexSlice (buf, start, end) {
  var len = buf.length

  if (!start || start < 0) start = 0
  if (!end || end < 0 || end > len) end = len

  var out = ''
  for (var i = start; i < end; ++i) {
    out += toHex(buf[i])
  }
  return out
}

function utf16leSlice (buf, start, end) {
  var bytes = buf.slice(start, end)
  var res = ''
  for (var i = 0; i < bytes.length; i += 2) {
    res += String.fromCharCode(bytes[i] + bytes[i + 1] * 256)
  }
  return res
}

Buffer.prototype.slice = function slice (start, end) {
  var len = this.length
  start = ~~start
  end = end === undefined ? len : ~~end

  if (start < 0) {
    start += len
    if (start < 0) start = 0
  } else if (start > len) {
    start = len
  }

  if (end < 0) {
    end += len
    if (end < 0) end = 0
  } else if (end > len) {
    end = len
  }

  if (end < start) end = start

  var newBuf
  if (Buffer.TYPED_ARRAY_SUPPORT) {
    newBuf = this.subarray(start, end)
    newBuf.__proto__ = Buffer.prototype
  } else {
    var sliceLen = end - start
    newBuf = new Buffer(sliceLen, undefined)
    for (var i = 0; i < sliceLen; ++i) {
      newBuf[i] = this[i + start]
    }
  }

  return newBuf
}

/*
 * Need to make sure that buffer isn't trying to write out of bounds.
 */
function checkOffset (offset, ext, length) {
  if ((offset % 1) !== 0 || offset < 0) throw new RangeError('offset is not uint')
  if (offset + ext > length) throw new RangeError('Trying to access beyond buffer length')
}

Buffer.prototype.readUIntLE = function readUIntLE (offset, byteLength, noAssert) {
  offset = offset | 0
  byteLength = byteLength | 0
  if (!noAssert) checkOffset(offset, byteLength, this.length)

  var val = this[offset]
  var mul = 1
  var i = 0
  while (++i < byteLength && (mul *= 0x100)) {
    val += this[offset + i] * mul
  }

  return val
}

Buffer.prototype.readUIntBE = function readUIntBE (offset, byteLength, noAssert) {
  offset = offset | 0
  byteLength = byteLength | 0
  if (!noAssert) {
    checkOffset(offset, byteLength, this.length)
  }

  var val = this[offset + --byteLength]
  var mul = 1
  while (byteLength > 0 && (mul *= 0x100)) {
    val += this[offset + --byteLength] * mul
  }

  return val
}

Buffer.prototype.readUInt8 = function readUInt8 (offset, noAssert) {
  if (!noAssert) checkOffset(offset, 1, this.length)
  return this[offset]
}

Buffer.prototype.readUInt16LE = function readUInt16LE (offset, noAssert) {
  if (!noAssert) checkOffset(offset, 2, this.length)
  return this[offset] | (this[offset + 1] << 8)
}

Buffer.prototype.readUInt16BE = function readUInt16BE (offset, noAssert) {
  if (!noAssert) checkOffset(offset, 2, this.length)
  return (this[offset] << 8) | this[offset + 1]
}

Buffer.prototype.readUInt32LE = function readUInt32LE (offset, noAssert) {
  if (!noAssert) checkOffset(offset, 4, this.length)

  return ((this[offset]) |
      (this[offset + 1] << 8) |
      (this[offset + 2] << 16)) +
      (this[offset + 3] * 0x1000000)
}

Buffer.prototype.readUInt32BE = function readUInt32BE (offset, noAssert) {
  if (!noAssert) checkOffset(offset, 4, this.length)

  return (this[offset] * 0x1000000) +
    ((this[offset + 1] << 16) |
    (this[offset + 2] << 8) |
    this[offset + 3])
}

Buffer.prototype.readIntLE = function readIntLE (offset, byteLength, noAssert) {
  offset = offset | 0
  byteLength = byteLength | 0
  if (!noAssert) checkOffset(offset, byteLength, this.length)

  var val = this[offset]
  var mul = 1
  var i = 0
  while (++i < byteLength && (mul *= 0x100)) {
    val += this[offset + i] * mul
  }
  mul *= 0x80

  if (val >= mul) val -= Math.pow(2, 8 * byteLength)

  return val
}

Buffer.prototype.readIntBE = function readIntBE (offset, byteLength, noAssert) {
  offset = offset | 0
  byteLength = byteLength | 0
  if (!noAssert) checkOffset(offset, byteLength, this.length)

  var i = byteLength
  var mul = 1
  var val = this[offset + --i]
  while (i > 0 && (mul *= 0x100)) {
    val += this[offset + --i] * mul
  }
  mul *= 0x80

  if (val >= mul) val -= Math.pow(2, 8 * byteLength)

  return val
}

Buffer.prototype.readInt8 = function readInt8 (offset, noAssert) {
  if (!noAssert) checkOffset(offset, 1, this.length)
  if (!(this[offset] & 0x80)) return (this[offset])
  return ((0xff - this[offset] + 1) * -1)
}

Buffer.prototype.readInt16LE = function readInt16LE (offset, noAssert) {
  if (!noAssert) checkOffset(offset, 2, this.length)
  var val = this[offset] | (this[offset + 1] << 8)
  return (val & 0x8000) ? val | 0xFFFF0000 : val
}

Buffer.prototype.readInt16BE = function readInt16BE (offset, noAssert) {
  if (!noAssert) checkOffset(offset, 2, this.length)
  var val = this[offset + 1] | (this[offset] << 8)
  return (val & 0x8000) ? val | 0xFFFF0000 : val
}

Buffer.prototype.readInt32LE = function readInt32LE (offset, noAssert) {
  if (!noAssert) checkOffset(offset, 4, this.length)

  return (this[offset]) |
    (this[offset + 1] << 8) |
    (this[offset + 2] << 16) |
    (this[offset + 3] << 24)
}

Buffer.prototype.readInt32BE = function readInt32BE (offset, noAssert) {
  if (!noAssert) checkOffset(offset, 4, this.length)

  return (this[offset] << 24) |
    (this[offset + 1] << 16) |
    (this[offset + 2] << 8) |
    (this[offset + 3])
}

Buffer.prototype.readFloatLE = function readFloatLE (offset, noAssert) {
  if (!noAssert) checkOffset(offset, 4, this.length)
  return ieee754.read(this, offset, true, 23, 4)
}

Buffer.prototype.readFloatBE = function readFloatBE (offset, noAssert) {
  if (!noAssert) checkOffset(offset, 4, this.length)
  return ieee754.read(this, offset, false, 23, 4)
}

Buffer.prototype.readDoubleLE = function readDoubleLE (offset, noAssert) {
  if (!noAssert) checkOffset(offset, 8, this.length)
  return ieee754.read(this, offset, true, 52, 8)
}

Buffer.prototype.readDoubleBE = function readDoubleBE (offset, noAssert) {
  if (!noAssert) checkOffset(offset, 8, this.length)
  return ieee754.read(this, offset, false, 52, 8)
}

function checkInt (buf, value, offset, ext, max, min) {
  if (!Buffer.isBuffer(buf)) throw new TypeError('"buffer" argument must be a Buffer instance')
  if (value > max || value < min) throw new RangeError('"value" argument is out of bounds')
  if (offset + ext > buf.length) throw new RangeError('Index out of range')
}

Buffer.prototype.writeUIntLE = function writeUIntLE (value, offset, byteLength, noAssert) {
  value = +value
  offset = offset | 0
  byteLength = byteLength | 0
  if (!noAssert) {
    var maxBytes = Math.pow(2, 8 * byteLength) - 1
    checkInt(this, value, offset, byteLength, maxBytes, 0)
  }

  var mul = 1
  var i = 0
  this[offset] = value & 0xFF
  while (++i < byteLength && (mul *= 0x100)) {
    this[offset + i] = (value / mul) & 0xFF
  }

  return offset + byteLength
}

Buffer.prototype.writeUIntBE = function writeUIntBE (value, offset, byteLength, noAssert) {
  value = +value
  offset = offset | 0
  byteLength = byteLength | 0
  if (!noAssert) {
    var maxBytes = Math.pow(2, 8 * byteLength) - 1
    checkInt(this, value, offset, byteLength, maxBytes, 0)
  }

  var i = byteLength - 1
  var mul = 1
  this[offset + i] = value & 0xFF
  while (--i >= 0 && (mul *= 0x100)) {
    this[offset + i] = (value / mul) & 0xFF
  }

  return offset + byteLength
}

Buffer.prototype.writeUInt8 = function writeUInt8 (value, offset, noAssert) {
  value = +value
  offset = offset | 0
  if (!noAssert) checkInt(this, value, offset, 1, 0xff, 0)
  if (!Buffer.TYPED_ARRAY_SUPPORT) value = Math.floor(value)
  this[offset] = (value & 0xff)
  return offset + 1
}

function objectWriteUInt16 (buf, value, offset, littleEndian) {
  if (value < 0) value = 0xffff + value + 1
  for (var i = 0, j = Math.min(buf.length - offset, 2); i < j; ++i) {
    buf[offset + i] = (value & (0xff << (8 * (littleEndian ? i : 1 - i)))) >>>
      (littleEndian ? i : 1 - i) * 8
  }
}

Buffer.prototype.writeUInt16LE = function writeUInt16LE (value, offset, noAssert) {
  value = +value
  offset = offset | 0
  if (!noAssert) checkInt(this, value, offset, 2, 0xffff, 0)
  if (Buffer.TYPED_ARRAY_SUPPORT) {
    this[offset] = (value & 0xff)
    this[offset + 1] = (value >>> 8)
  } else {
    objectWriteUInt16(this, value, offset, true)
  }
  return offset + 2
}

Buffer.prototype.writeUInt16BE = function writeUInt16BE (value, offset, noAssert) {
  value = +value
  offset = offset | 0
  if (!noAssert) checkInt(this, value, offset, 2, 0xffff, 0)
  if (Buffer.TYPED_ARRAY_SUPPORT) {
    this[offset] = (value >>> 8)
    this[offset + 1] = (value & 0xff)
  } else {
    objectWriteUInt16(this, value, offset, false)
  }
  return offset + 2
}

function objectWriteUInt32 (buf, value, offset, littleEndian) {
  if (value < 0) value = 0xffffffff + value + 1
  for (var i = 0, j = Math.min(buf.length - offset, 4); i < j; ++i) {
    buf[offset + i] = (value >>> (littleEndian ? i : 3 - i) * 8) & 0xff
  }
}

Buffer.prototype.writeUInt32LE = function writeUInt32LE (value, offset, noAssert) {
  value = +value
  offset = offset | 0
  if (!noAssert) checkInt(this, value, offset, 4, 0xffffffff, 0)
  if (Buffer.TYPED_ARRAY_SUPPORT) {
    this[offset + 3] = (value >>> 24)
    this[offset + 2] = (value >>> 16)
    this[offset + 1] = (value >>> 8)
    this[offset] = (value & 0xff)
  } else {
    objectWriteUInt32(this, value, offset, true)
  }
  return offset + 4
}

Buffer.prototype.writeUInt32BE = function writeUInt32BE (value, offset, noAssert) {
  value = +value
  offset = offset | 0
  if (!noAssert) checkInt(this, value, offset, 4, 0xffffffff, 0)
  if (Buffer.TYPED_ARRAY_SUPPORT) {
    this[offset] = (value >>> 24)
    this[offset + 1] = (value >>> 16)
    this[offset + 2] = (value >>> 8)
    this[offset + 3] = (value & 0xff)
  } else {
    objectWriteUInt32(this, value, offset, false)
  }
  return offset + 4
}

Buffer.prototype.writeIntLE = function writeIntLE (value, offset, byteLength, noAssert) {
  value = +value
  offset = offset | 0
  if (!noAssert) {
    var limit = Math.pow(2, 8 * byteLength - 1)

    checkInt(this, value, offset, byteLength, limit - 1, -limit)
  }

  var i = 0
  var mul = 1
  var sub = 0
  this[offset] = value & 0xFF
  while (++i < byteLength && (mul *= 0x100)) {
    if (value < 0 && sub === 0 && this[offset + i - 1] !== 0) {
      sub = 1
    }
    this[offset + i] = ((value / mul) >> 0) - sub & 0xFF
  }

  return offset + byteLength
}

Buffer.prototype.writeIntBE = function writeIntBE (value, offset, byteLength, noAssert) {
  value = +value
  offset = offset | 0
  if (!noAssert) {
    var limit = Math.pow(2, 8 * byteLength - 1)

    checkInt(this, value, offset, byteLength, limit - 1, -limit)
  }

  var i = byteLength - 1
  var mul = 1
  var sub = 0
  this[offset + i] = value & 0xFF
  while (--i >= 0 && (mul *= 0x100)) {
    if (value < 0 && sub === 0 && this[offset + i + 1] !== 0) {
      sub = 1
    }
    this[offset + i] = ((value / mul) >> 0) - sub & 0xFF
  }

  return offset + byteLength
}

Buffer.prototype.writeInt8 = function writeInt8 (value, offset, noAssert) {
  value = +value
  offset = offset | 0
  if (!noAssert) checkInt(this, value, offset, 1, 0x7f, -0x80)
  if (!Buffer.TYPED_ARRAY_SUPPORT) value = Math.floor(value)
  if (value < 0) value = 0xff + value + 1
  this[offset] = (value & 0xff)
  return offset + 1
}

Buffer.prototype.writeInt16LE = function writeInt16LE (value, offset, noAssert) {
  value = +value
  offset = offset | 0
  if (!noAssert) checkInt(this, value, offset, 2, 0x7fff, -0x8000)
  if (Buffer.TYPED_ARRAY_SUPPORT) {
    this[offset] = (value & 0xff)
    this[offset + 1] = (value >>> 8)
  } else {
    objectWriteUInt16(this, value, offset, true)
  }
  return offset + 2
}

Buffer.prototype.writeInt16BE = function writeInt16BE (value, offset, noAssert) {
  value = +value
  offset = offset | 0
  if (!noAssert) checkInt(this, value, offset, 2, 0x7fff, -0x8000)
  if (Buffer.TYPED_ARRAY_SUPPORT) {
    this[offset] = (value >>> 8)
    this[offset + 1] = (value & 0xff)
  } else {
    objectWriteUInt16(this, value, offset, false)
  }
  return offset + 2
}

Buffer.prototype.writeInt32LE = function writeInt32LE (value, offset, noAssert) {
  value = +value
  offset = offset | 0
  if (!noAssert) checkInt(this, value, offset, 4, 0x7fffffff, -0x80000000)
  if (Buffer.TYPED_ARRAY_SUPPORT) {
    this[offset] = (value & 0xff)
    this[offset + 1] = (value >>> 8)
    this[offset + 2] = (value >>> 16)
    this[offset + 3] = (value >>> 24)
  } else {
    objectWriteUInt32(this, value, offset, true)
  }
  return offset + 4
}

Buffer.prototype.writeInt32BE = function writeInt32BE (value, offset, noAssert) {
  value = +value
  offset = offset | 0
  if (!noAssert) checkInt(this, value, offset, 4, 0x7fffffff, -0x80000000)
  if (value < 0) value = 0xffffffff + value + 1
  if (Buffer.TYPED_ARRAY_SUPPORT) {
    this[offset] = (value >>> 24)
    this[offset + 1] = (value >>> 16)
    this[offset + 2] = (value >>> 8)
    this[offset + 3] = (value & 0xff)
  } else {
    objectWriteUInt32(this, value, offset, false)
  }
  return offset + 4
}

function checkIEEE754 (buf, value, offset, ext, max, min) {
  if (offset + ext > buf.length) throw new RangeError('Index out of range')
  if (offset < 0) throw new RangeError('Index out of range')
}

function writeFloat (buf, value, offset, littleEndian, noAssert) {
  if (!noAssert) {
    checkIEEE754(buf, value, offset, 4, 3.4028234663852886e+38, -3.4028234663852886e+38)
  }
  ieee754.write(buf, value, offset, littleEndian, 23, 4)
  return offset + 4
}

Buffer.prototype.writeFloatLE = function writeFloatLE (value, offset, noAssert) {
  return writeFloat(this, value, offset, true, noAssert)
}

Buffer.prototype.writeFloatBE = function writeFloatBE (value, offset, noAssert) {
  return writeFloat(this, value, offset, false, noAssert)
}

function writeDouble (buf, value, offset, littleEndian, noAssert) {
  if (!noAssert) {
    checkIEEE754(buf, value, offset, 8, 1.7976931348623157E+308, -1.7976931348623157E+308)
  }
  ieee754.write(buf, value, offset, littleEndian, 52, 8)
  return offset + 8
}

Buffer.prototype.writeDoubleLE = function writeDoubleLE (value, offset, noAssert) {
  return writeDouble(this, value, offset, true, noAssert)
}

Buffer.prototype.writeDoubleBE = function writeDoubleBE (value, offset, noAssert) {
  return writeDouble(this, value, offset, false, noAssert)
}

// copy(targetBuffer, targetStart=0, sourceStart=0, sourceEnd=buffer.length)
Buffer.prototype.copy = function copy (target, targetStart, start, end) {
  if (!start) start = 0
  if (!end && end !== 0) end = this.length
  if (targetStart >= target.length) targetStart = target.length
  if (!targetStart) targetStart = 0
  if (end > 0 && end < start) end = start

  // Copy 0 bytes; we're done
  if (end === start) return 0
  if (target.length === 0 || this.length === 0) return 0

  // Fatal error conditions
  if (targetStart < 0) {
    throw new RangeError('targetStart out of bounds')
  }
  if (start < 0 || start >= this.length) throw new RangeError('sourceStart out of bounds')
  if (end < 0) throw new RangeError('sourceEnd out of bounds')

  // Are we oob?
  if (end > this.length) end = this.length
  if (target.length - targetStart < end - start) {
    end = target.length - targetStart + start
  }

  var len = end - start
  var i

  if (this === target && start < targetStart && targetStart < end) {
    // descending copy from end
    for (i = len - 1; i >= 0; --i) {
      target[i + targetStart] = this[i + start]
    }
  } else if (len < 1000 || !Buffer.TYPED_ARRAY_SUPPORT) {
    // ascending copy from start
    for (i = 0; i < len; ++i) {
      target[i + targetStart] = this[i + start]
    }
  } else {
    Uint8Array.prototype.set.call(
      target,
      this.subarray(start, start + len),
      targetStart
    )
  }

  return len
}

// Usage:
//    buffer.fill(number[, offset[, end]])
//    buffer.fill(buffer[, offset[, end]])
//    buffer.fill(string[, offset[, end]][, encoding])
Buffer.prototype.fill = function fill (val, start, end, encoding) {
  // Handle string cases:
  if (typeof val === 'string') {
    if (typeof start === 'string') {
      encoding = start
      start = 0
      end = this.length
    } else if (typeof end === 'string') {
      encoding = end
      end = this.length
    }
    if (val.length === 1) {
      var code = val.charCodeAt(0)
      if (code < 256) {
        val = code
      }
    }
    if (encoding !== undefined && typeof encoding !== 'string') {
      throw new TypeError('encoding must be a string')
    }
    if (typeof encoding === 'string' && !Buffer.isEncoding(encoding)) {
      throw new TypeError('Unknown encoding: ' + encoding)
    }
  } else if (typeof val === 'number') {
    val = val & 255
  }

  // Invalid ranges are not set to a default, so can range check early.
  if (start < 0 || this.length < start || this.length < end) {
    throw new RangeError('Out of range index')
  }

  if (end <= start) {
    return this
  }

  start = start >>> 0
  end = end === undefined ? this.length : end >>> 0

  if (!val) val = 0

  var i
  if (typeof val === 'number') {
    for (i = start; i < end; ++i) {
      this[i] = val
    }
  } else {
    var bytes = Buffer.isBuffer(val)
      ? val
      : utf8ToBytes(new Buffer(val, encoding).toString())
    var len = bytes.length
    for (i = 0; i < end - start; ++i) {
      this[i + start] = bytes[i % len]
    }
  }

  return this
}

// HELPER FUNCTIONS
// ================

var INVALID_BASE64_RE = /[^+\/0-9A-Za-z-_]/g

function base64clean (str) {
  // Node strips out invalid characters like \n and \t from the string, base64-js does not
  str = stringtrim(str).replace(INVALID_BASE64_RE, '')
  // Node converts strings with length < 2 to ''
  if (str.length < 2) return ''
  // Node allows for non-padded base64 strings (missing trailing ===), base64-js does not
  while (str.length % 4 !== 0) {
    str = str + '='
  }
  return str
}

function stringtrim (str) {
  if (str.trim) return str.trim()
  return str.replace(/^\s+|\s+$/g, '')
}

function toHex (n) {
  if (n < 16) return '0' + n.toString(16)
  return n.toString(16)
}

function utf8ToBytes (string, units) {
  units = units || Infinity
  var codePoint
  var length = string.length
  var leadSurrogate = null
  var bytes = []

  for (var i = 0; i < length; ++i) {
    codePoint = string.charCodeAt(i)

    // is surrogate component
    if (codePoint > 0xD7FF && codePoint < 0xE000) {
      // last char was a lead
      if (!leadSurrogate) {
        // no lead yet
        if (codePoint > 0xDBFF) {
          // unexpected trail
          if ((units -= 3) > -1) bytes.push(0xEF, 0xBF, 0xBD)
          continue
        } else if (i + 1 === length) {
          // unpaired lead
          if ((units -= 3) > -1) bytes.push(0xEF, 0xBF, 0xBD)
          continue
        }

        // valid lead
        leadSurrogate = codePoint

        continue
      }

      // 2 leads in a row
      if (codePoint < 0xDC00) {
        if ((units -= 3) > -1) bytes.push(0xEF, 0xBF, 0xBD)
        leadSurrogate = codePoint
        continue
      }

      // valid surrogate pair
      codePoint = (leadSurrogate - 0xD800 << 10 | codePoint - 0xDC00) + 0x10000
    } else if (leadSurrogate) {
      // valid bmp char, but last char was a lead
      if ((units -= 3) > -1) bytes.push(0xEF, 0xBF, 0xBD)
    }

    leadSurrogate = null

    // encode utf8
    if (codePoint < 0x80) {
      if ((units -= 1) < 0) break
      bytes.push(codePoint)
    } else if (codePoint < 0x800) {
      if ((units -= 2) < 0) break
      bytes.push(
        codePoint >> 0x6 | 0xC0,
        codePoint & 0x3F | 0x80
      )
    } else if (codePoint < 0x10000) {
      if ((units -= 3) < 0) break
      bytes.push(
        codePoint >> 0xC | 0xE0,
        codePoint >> 0x6 & 0x3F | 0x80,
        codePoint & 0x3F | 0x80
      )
    } else if (codePoint < 0x110000) {
      if ((units -= 4) < 0) break
      bytes.push(
        codePoint >> 0x12 | 0xF0,
        codePoint >> 0xC & 0x3F | 0x80,
        codePoint >> 0x6 & 0x3F | 0x80,
        codePoint & 0x3F | 0x80
      )
    } else {
      throw new Error('Invalid code point')
    }
  }

  return bytes
}

function asciiToBytes (str) {
  var byteArray = []
  for (var i = 0; i < str.length; ++i) {
    // Node's code seems to be doing this and not & 0x7F..
    byteArray.push(str.charCodeAt(i) & 0xFF)
  }
  return byteArray
}

function utf16leToBytes (str, units) {
  var c, hi, lo
  var byteArray = []
  for (var i = 0; i < str.length; ++i) {
    if ((units -= 2) < 0) break

    c = str.charCodeAt(i)
    hi = c >> 8
    lo = c % 256
    byteArray.push(lo)
    byteArray.push(hi)
  }

  return byteArray
}

function base64ToBytes (str) {
  return base64.toByteArray(base64clean(str))
}

function blitBuffer (src, dst, offset, length) {
  for (var i = 0; i < length; ++i) {
    if ((i + offset >= dst.length) || (i >= src.length)) break
    dst[i + offset] = src[i]
  }
  return i
}

function isnan (val) {
  return val !== val // eslint-disable-line no-self-compare
}

}).call(this)}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {},require("buffer").Buffer)
},{"base64-js":1,"buffer":2,"ieee754":4,"isarray":5}],3:[function(require,module,exports){
!function(t,e){"object"==typeof exports&&"undefined"!=typeof module?module.exports=e():"function"==typeof define&&define.amd?define(e):(t="undefined"!=typeof globalThis?globalThis:t||self).dayjs=e()}(this,(function(){"use strict";var t=1e3,e=6e4,n=36e5,r="millisecond",i="second",s="minute",u="hour",a="day",o="week",f="month",h="quarter",c="year",d="date",$="Invalid Date",l=/^(\d{4})[-/]?(\d{1,2})?[-/]?(\d{0,2})[Tt\s]*(\d{1,2})?:?(\d{1,2})?:?(\d{1,2})?[.:]?(\d+)?$/,y=/\[([^\]]+)]|Y{1,4}|M{1,4}|D{1,2}|d{1,4}|H{1,2}|h{1,2}|a|A|m{1,2}|s{1,2}|Z{1,2}|SSS/g,M={name:"en",weekdays:"Sunday_Monday_Tuesday_Wednesday_Thursday_Friday_Saturday".split("_"),months:"January_February_March_April_May_June_July_August_September_October_November_December".split("_")},m=function(t,e,n){var r=String(t);return!r||r.length>=e?t:""+Array(e+1-r.length).join(n)+t},g={s:m,z:function(t){var e=-t.utcOffset(),n=Math.abs(e),r=Math.floor(n/60),i=n%60;return(e<=0?"+":"-")+m(r,2,"0")+":"+m(i,2,"0")},m:function t(e,n){if(e.date()<n.date())return-t(n,e);var r=12*(n.year()-e.year())+(n.month()-e.month()),i=e.clone().add(r,f),s=n-i<0,u=e.clone().add(r+(s?-1:1),f);return+(-(r+(n-i)/(s?i-u:u-i))||0)},a:function(t){return t<0?Math.ceil(t)||0:Math.floor(t)},p:function(t){return{M:f,y:c,w:o,d:a,D:d,h:u,m:s,s:i,ms:r,Q:h}[t]||String(t||"").toLowerCase().replace(/s$/,"")},u:function(t){return void 0===t}},D="en",v={};v[D]=M;var p=function(t){return t instanceof _},S=function(t,e,n){var r;if(!t)return D;if("string"==typeof t)v[t]&&(r=t),e&&(v[t]=e,r=t);else{var i=t.name;v[i]=t,r=i}return!n&&r&&(D=r),r||!n&&D},w=function(t,e){if(p(t))return t.clone();var n="object"==typeof e?e:{};return n.date=t,n.args=arguments,new _(n)},O=g;O.l=S,O.i=p,O.w=function(t,e){return w(t,{locale:e.$L,utc:e.$u,x:e.$x,$offset:e.$offset})};var _=function(){function M(t){this.$L=S(t.locale,null,!0),this.parse(t)}var m=M.prototype;return m.parse=function(t){this.$d=function(t){var e=t.date,n=t.utc;if(null===e)return new Date(NaN);if(O.u(e))return new Date;if(e instanceof Date)return new Date(e);if("string"==typeof e&&!/Z$/i.test(e)){var r=e.match(l);if(r){var i=r[2]-1||0,s=(r[7]||"0").substring(0,3);return n?new Date(Date.UTC(r[1],i,r[3]||1,r[4]||0,r[5]||0,r[6]||0,s)):new Date(r[1],i,r[3]||1,r[4]||0,r[5]||0,r[6]||0,s)}}return new Date(e)}(t),this.$x=t.x||{},this.init()},m.init=function(){var t=this.$d;this.$y=t.getFullYear(),this.$M=t.getMonth(),this.$D=t.getDate(),this.$W=t.getDay(),this.$H=t.getHours(),this.$m=t.getMinutes(),this.$s=t.getSeconds(),this.$ms=t.getMilliseconds()},m.$utils=function(){return O},m.isValid=function(){return!(this.$d.toString()===$)},m.isSame=function(t,e){var n=w(t);return this.startOf(e)<=n&&n<=this.endOf(e)},m.isAfter=function(t,e){return w(t)<this.startOf(e)},m.isBefore=function(t,e){return this.endOf(e)<w(t)},m.$g=function(t,e,n){return O.u(t)?this[e]:this.set(n,t)},m.unix=function(){return Math.floor(this.valueOf()/1e3)},m.valueOf=function(){return this.$d.getTime()},m.startOf=function(t,e){var n=this,r=!!O.u(e)||e,h=O.p(t),$=function(t,e){var i=O.w(n.$u?Date.UTC(n.$y,e,t):new Date(n.$y,e,t),n);return r?i:i.endOf(a)},l=function(t,e){return O.w(n.toDate()[t].apply(n.toDate("s"),(r?[0,0,0,0]:[23,59,59,999]).slice(e)),n)},y=this.$W,M=this.$M,m=this.$D,g="set"+(this.$u?"UTC":"");switch(h){case c:return r?$(1,0):$(31,11);case f:return r?$(1,M):$(0,M+1);case o:var D=this.$locale().weekStart||0,v=(y<D?y+7:y)-D;return $(r?m-v:m+(6-v),M);case a:case d:return l(g+"Hours",0);case u:return l(g+"Minutes",1);case s:return l(g+"Seconds",2);case i:return l(g+"Milliseconds",3);default:return this.clone()}},m.endOf=function(t){return this.startOf(t,!1)},m.$set=function(t,e){var n,o=O.p(t),h="set"+(this.$u?"UTC":""),$=(n={},n[a]=h+"Date",n[d]=h+"Date",n[f]=h+"Month",n[c]=h+"FullYear",n[u]=h+"Hours",n[s]=h+"Minutes",n[i]=h+"Seconds",n[r]=h+"Milliseconds",n)[o],l=o===a?this.$D+(e-this.$W):e;if(o===f||o===c){var y=this.clone().set(d,1);y.$d[$](l),y.init(),this.$d=y.set(d,Math.min(this.$D,y.daysInMonth())).$d}else $&&this.$d[$](l);return this.init(),this},m.set=function(t,e){return this.clone().$set(t,e)},m.get=function(t){return this[O.p(t)]()},m.add=function(r,h){var d,$=this;r=Number(r);var l=O.p(h),y=function(t){var e=w($);return O.w(e.date(e.date()+Math.round(t*r)),$)};if(l===f)return this.set(f,this.$M+r);if(l===c)return this.set(c,this.$y+r);if(l===a)return y(1);if(l===o)return y(7);var M=(d={},d[s]=e,d[u]=n,d[i]=t,d)[l]||1,m=this.$d.getTime()+r*M;return O.w(m,this)},m.subtract=function(t,e){return this.add(-1*t,e)},m.format=function(t){var e=this,n=this.$locale();if(!this.isValid())return n.invalidDate||$;var r=t||"YYYY-MM-DDTHH:mm:ssZ",i=O.z(this),s=this.$H,u=this.$m,a=this.$M,o=n.weekdays,f=n.months,h=function(t,n,i,s){return t&&(t[n]||t(e,r))||i[n].substr(0,s)},c=function(t){return O.s(s%12||12,t,"0")},d=n.meridiem||function(t,e,n){var r=t<12?"AM":"PM";return n?r.toLowerCase():r},l={YY:String(this.$y).slice(-2),YYYY:this.$y,M:a+1,MM:O.s(a+1,2,"0"),MMM:h(n.monthsShort,a,f,3),MMMM:h(f,a),D:this.$D,DD:O.s(this.$D,2,"0"),d:String(this.$W),dd:h(n.weekdaysMin,this.$W,o,2),ddd:h(n.weekdaysShort,this.$W,o,3),dddd:o[this.$W],H:String(s),HH:O.s(s,2,"0"),h:c(1),hh:c(2),a:d(s,u,!0),A:d(s,u,!1),m:String(u),mm:O.s(u,2,"0"),s:String(this.$s),ss:O.s(this.$s,2,"0"),SSS:O.s(this.$ms,3,"0"),Z:i};return r.replace(y,(function(t,e){return e||l[t]||i.replace(":","")}))},m.utcOffset=function(){return 15*-Math.round(this.$d.getTimezoneOffset()/15)},m.diff=function(r,d,$){var l,y=O.p(d),M=w(r),m=(M.utcOffset()-this.utcOffset())*e,g=this-M,D=O.m(this,M);return D=(l={},l[c]=D/12,l[f]=D,l[h]=D/3,l[o]=(g-m)/6048e5,l[a]=(g-m)/864e5,l[u]=g/n,l[s]=g/e,l[i]=g/t,l)[y]||g,$?D:O.a(D)},m.daysInMonth=function(){return this.endOf(f).$D},m.$locale=function(){return v[this.$L]},m.locale=function(t,e){if(!t)return this.$L;var n=this.clone(),r=S(t,e,!0);return r&&(n.$L=r),n},m.clone=function(){return O.w(this.$d,this)},m.toDate=function(){return new Date(this.valueOf())},m.toJSON=function(){return this.isValid()?this.toISOString():null},m.toISOString=function(){return this.$d.toISOString()},m.toString=function(){return this.$d.toUTCString()},M}(),b=_.prototype;return w.prototype=b,[["$ms",r],["$s",i],["$m",s],["$H",u],["$W",a],["$M",f],["$y",c],["$D",d]].forEach((function(t){b[t[1]]=function(e){return this.$g(e,t[0],t[1])}})),w.extend=function(t,e){return t.$i||(t(e,_,w),t.$i=!0),w},w.locale=S,w.isDayjs=p,w.unix=function(t){return w(1e3*t)},w.en=v[D],w.Ls=v,w.p={},w}));
},{}],4:[function(require,module,exports){
/*! ieee754. BSD-3-Clause License. Feross Aboukhadijeh <https://feross.org/opensource> */
exports.read = function (buffer, offset, isLE, mLen, nBytes) {
  var e, m
  var eLen = (nBytes * 8) - mLen - 1
  var eMax = (1 << eLen) - 1
  var eBias = eMax >> 1
  var nBits = -7
  var i = isLE ? (nBytes - 1) : 0
  var d = isLE ? -1 : 1
  var s = buffer[offset + i]

  i += d

  e = s & ((1 << (-nBits)) - 1)
  s >>= (-nBits)
  nBits += eLen
  for (; nBits > 0; e = (e * 256) + buffer[offset + i], i += d, nBits -= 8) {}

  m = e & ((1 << (-nBits)) - 1)
  e >>= (-nBits)
  nBits += mLen
  for (; nBits > 0; m = (m * 256) + buffer[offset + i], i += d, nBits -= 8) {}

  if (e === 0) {
    e = 1 - eBias
  } else if (e === eMax) {
    return m ? NaN : ((s ? -1 : 1) * Infinity)
  } else {
    m = m + Math.pow(2, mLen)
    e = e - eBias
  }
  return (s ? -1 : 1) * m * Math.pow(2, e - mLen)
}

exports.write = function (buffer, value, offset, isLE, mLen, nBytes) {
  var e, m, c
  var eLen = (nBytes * 8) - mLen - 1
  var eMax = (1 << eLen) - 1
  var eBias = eMax >> 1
  var rt = (mLen === 23 ? Math.pow(2, -24) - Math.pow(2, -77) : 0)
  var i = isLE ? 0 : (nBytes - 1)
  var d = isLE ? 1 : -1
  var s = value < 0 || (value === 0 && 1 / value < 0) ? 1 : 0

  value = Math.abs(value)

  if (isNaN(value) || value === Infinity) {
    m = isNaN(value) ? 1 : 0
    e = eMax
  } else {
    e = Math.floor(Math.log(value) / Math.LN2)
    if (value * (c = Math.pow(2, -e)) < 1) {
      e--
      c *= 2
    }
    if (e + eBias >= 1) {
      value += rt / c
    } else {
      value += rt * Math.pow(2, 1 - eBias)
    }
    if (value * c >= 2) {
      e++
      c /= 2
    }

    if (e + eBias >= eMax) {
      m = 0
      e = eMax
    } else if (e + eBias >= 1) {
      m = ((value * c) - 1) * Math.pow(2, mLen)
      e = e + eBias
    } else {
      m = value * Math.pow(2, eBias - 1) * Math.pow(2, mLen)
      e = 0
    }
  }

  for (; mLen >= 8; buffer[offset + i] = m & 0xff, i += d, m /= 256, mLen -= 8) {}

  e = (e << mLen) | m
  eLen += mLen
  for (; eLen > 0; buffer[offset + i] = e & 0xff, i += d, e /= 256, eLen -= 8) {}

  buffer[offset + i - d] |= s * 128
}

},{}],5:[function(require,module,exports){
var toString = {}.toString;

module.exports = Array.isArray || function (arr) {
  return toString.call(arr) == '[object Array]';
};

},{}],6:[function(require,module,exports){
(function (global,Buffer,setImmediate){(function (){
/*!

JSZip v3.7.1 - A JavaScript class for generating and reading zip files
<http://stuartk.com/jszip>

(c) 2009-2016 Stuart Knightley <stuart [at] stuartk.com>
Dual licenced under the MIT license or GPLv3. See https://raw.github.com/Stuk/jszip/master/LICENSE.markdown.

JSZip uses the library pako released under the MIT license :
https://github.com/nodeca/pako/blob/master/LICENSE
*/

!function(t){if("object"==typeof exports&&"undefined"!=typeof module)module.exports=t();else if("function"==typeof define&&define.amd)define([],t);else{("undefined"!=typeof window?window:"undefined"!=typeof global?global:"undefined"!=typeof self?self:this).JSZip=t()}}(function(){return function s(a,o,h){function u(r,t){if(!o[r]){if(!a[r]){var e="function"==typeof require&&require;if(!t&&e)return e(r,!0);if(l)return l(r,!0);var i=new Error("Cannot find module '"+r+"'");throw i.code="MODULE_NOT_FOUND",i}var n=o[r]={exports:{}};a[r][0].call(n.exports,function(t){var e=a[r][1][t];return u(e||t)},n,n.exports,s,a,o,h)}return o[r].exports}for(var l="function"==typeof require&&require,t=0;t<h.length;t++)u(h[t]);return u}({1:[function(t,e,r){"use strict";var c=t("./utils"),d=t("./support"),p="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";r.encode=function(t){for(var e,r,i,n,s,a,o,h=[],u=0,l=t.length,f=l,d="string"!==c.getTypeOf(t);u<t.length;)f=l-u,i=d?(e=t[u++],r=u<l?t[u++]:0,u<l?t[u++]:0):(e=t.charCodeAt(u++),r=u<l?t.charCodeAt(u++):0,u<l?t.charCodeAt(u++):0),n=e>>2,s=(3&e)<<4|r>>4,a=1<f?(15&r)<<2|i>>6:64,o=2<f?63&i:64,h.push(p.charAt(n)+p.charAt(s)+p.charAt(a)+p.charAt(o));return h.join("")},r.decode=function(t){var e,r,i,n,s,a,o=0,h=0,u="data:";if(t.substr(0,u.length)===u)throw new Error("Invalid base64 input, it looks like a data url.");var l,f=3*(t=t.replace(/[^A-Za-z0-9\+\/\=]/g,"")).length/4;if(t.charAt(t.length-1)===p.charAt(64)&&f--,t.charAt(t.length-2)===p.charAt(64)&&f--,f%1!=0)throw new Error("Invalid base64 input, bad content length.");for(l=d.uint8array?new Uint8Array(0|f):new Array(0|f);o<t.length;)e=p.indexOf(t.charAt(o++))<<2|(n=p.indexOf(t.charAt(o++)))>>4,r=(15&n)<<4|(s=p.indexOf(t.charAt(o++)))>>2,i=(3&s)<<6|(a=p.indexOf(t.charAt(o++))),l[h++]=e,64!==s&&(l[h++]=r),64!==a&&(l[h++]=i);return l}},{"./support":30,"./utils":32}],2:[function(t,e,r){"use strict";var i=t("./external"),n=t("./stream/DataWorker"),s=t("./stream/Crc32Probe"),a=t("./stream/DataLengthProbe");function o(t,e,r,i,n){this.compressedSize=t,this.uncompressedSize=e,this.crc32=r,this.compression=i,this.compressedContent=n}o.prototype={getContentWorker:function(){var t=new n(i.Promise.resolve(this.compressedContent)).pipe(this.compression.uncompressWorker()).pipe(new a("data_length")),e=this;return t.on("end",function(){if(this.streamInfo.data_length!==e.uncompressedSize)throw new Error("Bug : uncompressed data size mismatch")}),t},getCompressedWorker:function(){return new n(i.Promise.resolve(this.compressedContent)).withStreamInfo("compressedSize",this.compressedSize).withStreamInfo("uncompressedSize",this.uncompressedSize).withStreamInfo("crc32",this.crc32).withStreamInfo("compression",this.compression)}},o.createWorkerFrom=function(t,e,r){return t.pipe(new s).pipe(new a("uncompressedSize")).pipe(e.compressWorker(r)).pipe(new a("compressedSize")).withStreamInfo("compression",e)},e.exports=o},{"./external":6,"./stream/Crc32Probe":25,"./stream/DataLengthProbe":26,"./stream/DataWorker":27}],3:[function(t,e,r){"use strict";var i=t("./stream/GenericWorker");r.STORE={magic:"\0\0",compressWorker:function(t){return new i("STORE compression")},uncompressWorker:function(){return new i("STORE decompression")}},r.DEFLATE=t("./flate")},{"./flate":7,"./stream/GenericWorker":28}],4:[function(t,e,r){"use strict";var i=t("./utils");var o=function(){for(var t,e=[],r=0;r<256;r++){t=r;for(var i=0;i<8;i++)t=1&t?3988292384^t>>>1:t>>>1;e[r]=t}return e}();e.exports=function(t,e){return void 0!==t&&t.length?"string"!==i.getTypeOf(t)?function(t,e,r,i){var n=o,s=i+r;t^=-1;for(var a=i;a<s;a++)t=t>>>8^n[255&(t^e[a])];return-1^t}(0|e,t,t.length,0):function(t,e,r,i){var n=o,s=i+r;t^=-1;for(var a=i;a<s;a++)t=t>>>8^n[255&(t^e.charCodeAt(a))];return-1^t}(0|e,t,t.length,0):0}},{"./utils":32}],5:[function(t,e,r){"use strict";r.base64=!1,r.binary=!1,r.dir=!1,r.createFolders=!0,r.date=null,r.compression=null,r.compressionOptions=null,r.comment=null,r.unixPermissions=null,r.dosPermissions=null},{}],6:[function(t,e,r){"use strict";var i=null;i="undefined"!=typeof Promise?Promise:t("lie"),e.exports={Promise:i}},{lie:37}],7:[function(t,e,r){"use strict";var i="undefined"!=typeof Uint8Array&&"undefined"!=typeof Uint16Array&&"undefined"!=typeof Uint32Array,n=t("pako"),s=t("./utils"),a=t("./stream/GenericWorker"),o=i?"uint8array":"array";function h(t,e){a.call(this,"FlateWorker/"+t),this._pako=null,this._pakoAction=t,this._pakoOptions=e,this.meta={}}r.magic="\b\0",s.inherits(h,a),h.prototype.processChunk=function(t){this.meta=t.meta,null===this._pako&&this._createPako(),this._pako.push(s.transformTo(o,t.data),!1)},h.prototype.flush=function(){a.prototype.flush.call(this),null===this._pako&&this._createPako(),this._pako.push([],!0)},h.prototype.cleanUp=function(){a.prototype.cleanUp.call(this),this._pako=null},h.prototype._createPako=function(){this._pako=new n[this._pakoAction]({raw:!0,level:this._pakoOptions.level||-1});var e=this;this._pako.onData=function(t){e.push({data:t,meta:e.meta})}},r.compressWorker=function(t){return new h("Deflate",t)},r.uncompressWorker=function(){return new h("Inflate",{})}},{"./stream/GenericWorker":28,"./utils":32,pako:38}],8:[function(t,e,r){"use strict";function A(t,e){var r,i="";for(r=0;r<e;r++)i+=String.fromCharCode(255&t),t>>>=8;return i}function i(t,e,r,i,n,s){var a,o,h=t.file,u=t.compression,l=s!==O.utf8encode,f=I.transformTo("string",s(h.name)),d=I.transformTo("string",O.utf8encode(h.name)),c=h.comment,p=I.transformTo("string",s(c)),m=I.transformTo("string",O.utf8encode(c)),_=d.length!==h.name.length,g=m.length!==c.length,b="",v="",y="",w=h.dir,k=h.date,x={crc32:0,compressedSize:0,uncompressedSize:0};e&&!r||(x.crc32=t.crc32,x.compressedSize=t.compressedSize,x.uncompressedSize=t.uncompressedSize);var S=0;e&&(S|=8),l||!_&&!g||(S|=2048);var z=0,C=0;w&&(z|=16),"UNIX"===n?(C=798,z|=function(t,e){var r=t;return t||(r=e?16893:33204),(65535&r)<<16}(h.unixPermissions,w)):(C=20,z|=function(t){return 63&(t||0)}(h.dosPermissions)),a=k.getUTCHours(),a<<=6,a|=k.getUTCMinutes(),a<<=5,a|=k.getUTCSeconds()/2,o=k.getUTCFullYear()-1980,o<<=4,o|=k.getUTCMonth()+1,o<<=5,o|=k.getUTCDate(),_&&(v=A(1,1)+A(B(f),4)+d,b+="up"+A(v.length,2)+v),g&&(y=A(1,1)+A(B(p),4)+m,b+="uc"+A(y.length,2)+y);var E="";return E+="\n\0",E+=A(S,2),E+=u.magic,E+=A(a,2),E+=A(o,2),E+=A(x.crc32,4),E+=A(x.compressedSize,4),E+=A(x.uncompressedSize,4),E+=A(f.length,2),E+=A(b.length,2),{fileRecord:R.LOCAL_FILE_HEADER+E+f+b,dirRecord:R.CENTRAL_FILE_HEADER+A(C,2)+E+A(p.length,2)+"\0\0\0\0"+A(z,4)+A(i,4)+f+b+p}}var I=t("../utils"),n=t("../stream/GenericWorker"),O=t("../utf8"),B=t("../crc32"),R=t("../signature");function s(t,e,r,i){n.call(this,"ZipFileWorker"),this.bytesWritten=0,this.zipComment=e,this.zipPlatform=r,this.encodeFileName=i,this.streamFiles=t,this.accumulate=!1,this.contentBuffer=[],this.dirRecords=[],this.currentSourceOffset=0,this.entriesCount=0,this.currentFile=null,this._sources=[]}I.inherits(s,n),s.prototype.push=function(t){var e=t.meta.percent||0,r=this.entriesCount,i=this._sources.length;this.accumulate?this.contentBuffer.push(t):(this.bytesWritten+=t.data.length,n.prototype.push.call(this,{data:t.data,meta:{currentFile:this.currentFile,percent:r?(e+100*(r-i-1))/r:100}}))},s.prototype.openedSource=function(t){this.currentSourceOffset=this.bytesWritten,this.currentFile=t.file.name;var e=this.streamFiles&&!t.file.dir;if(e){var r=i(t,e,!1,this.currentSourceOffset,this.zipPlatform,this.encodeFileName);this.push({data:r.fileRecord,meta:{percent:0}})}else this.accumulate=!0},s.prototype.closedSource=function(t){this.accumulate=!1;var e=this.streamFiles&&!t.file.dir,r=i(t,e,!0,this.currentSourceOffset,this.zipPlatform,this.encodeFileName);if(this.dirRecords.push(r.dirRecord),e)this.push({data:function(t){return R.DATA_DESCRIPTOR+A(t.crc32,4)+A(t.compressedSize,4)+A(t.uncompressedSize,4)}(t),meta:{percent:100}});else for(this.push({data:r.fileRecord,meta:{percent:0}});this.contentBuffer.length;)this.push(this.contentBuffer.shift());this.currentFile=null},s.prototype.flush=function(){for(var t=this.bytesWritten,e=0;e<this.dirRecords.length;e++)this.push({data:this.dirRecords[e],meta:{percent:100}});var r=this.bytesWritten-t,i=function(t,e,r,i,n){var s=I.transformTo("string",n(i));return R.CENTRAL_DIRECTORY_END+"\0\0\0\0"+A(t,2)+A(t,2)+A(e,4)+A(r,4)+A(s.length,2)+s}(this.dirRecords.length,r,t,this.zipComment,this.encodeFileName);this.push({data:i,meta:{percent:100}})},s.prototype.prepareNextSource=function(){this.previous=this._sources.shift(),this.openedSource(this.previous.streamInfo),this.isPaused?this.previous.pause():this.previous.resume()},s.prototype.registerPrevious=function(t){this._sources.push(t);var e=this;return t.on("data",function(t){e.processChunk(t)}),t.on("end",function(){e.closedSource(e.previous.streamInfo),e._sources.length?e.prepareNextSource():e.end()}),t.on("error",function(t){e.error(t)}),this},s.prototype.resume=function(){return!!n.prototype.resume.call(this)&&(!this.previous&&this._sources.length?(this.prepareNextSource(),!0):this.previous||this._sources.length||this.generatedError?void 0:(this.end(),!0))},s.prototype.error=function(t){var e=this._sources;if(!n.prototype.error.call(this,t))return!1;for(var r=0;r<e.length;r++)try{e[r].error(t)}catch(t){}return!0},s.prototype.lock=function(){n.prototype.lock.call(this);for(var t=this._sources,e=0;e<t.length;e++)t[e].lock()},e.exports=s},{"../crc32":4,"../signature":23,"../stream/GenericWorker":28,"../utf8":31,"../utils":32}],9:[function(t,e,r){"use strict";var u=t("../compressions"),i=t("./ZipFileWorker");r.generateWorker=function(t,a,e){var o=new i(a.streamFiles,e,a.platform,a.encodeFileName),h=0;try{t.forEach(function(t,e){h++;var r=function(t,e){var r=t||e,i=u[r];if(!i)throw new Error(r+" is not a valid compression method !");return i}(e.options.compression,a.compression),i=e.options.compressionOptions||a.compressionOptions||{},n=e.dir,s=e.date;e._compressWorker(r,i).withStreamInfo("file",{name:t,dir:n,date:s,comment:e.comment||"",unixPermissions:e.unixPermissions,dosPermissions:e.dosPermissions}).pipe(o)}),o.entriesCount=h}catch(t){o.error(t)}return o}},{"../compressions":3,"./ZipFileWorker":8}],10:[function(t,e,r){"use strict";function i(){if(!(this instanceof i))return new i;if(arguments.length)throw new Error("The constructor with parameters has been removed in JSZip 3.0, please check the upgrade guide.");this.files=Object.create(null),this.comment=null,this.root="",this.clone=function(){var t=new i;for(var e in this)"function"!=typeof this[e]&&(t[e]=this[e]);return t}}(i.prototype=t("./object")).loadAsync=t("./load"),i.support=t("./support"),i.defaults=t("./defaults"),i.version="3.7.1",i.loadAsync=function(t,e){return(new i).loadAsync(t,e)},i.external=t("./external"),e.exports=i},{"./defaults":5,"./external":6,"./load":11,"./object":15,"./support":30}],11:[function(t,e,r){"use strict";var i=t("./utils"),n=t("./external"),o=t("./utf8"),h=t("./zipEntries"),s=t("./stream/Crc32Probe"),u=t("./nodejsUtils");function l(i){return new n.Promise(function(t,e){var r=i.decompressed.getContentWorker().pipe(new s);r.on("error",function(t){e(t)}).on("end",function(){r.streamInfo.crc32!==i.decompressed.crc32?e(new Error("Corrupted zip : CRC32 mismatch")):t()}).resume()})}e.exports=function(t,s){var a=this;return s=i.extend(s||{},{base64:!1,checkCRC32:!1,optimizedBinaryString:!1,createFolders:!1,decodeFileName:o.utf8decode}),u.isNode&&u.isStream(t)?n.Promise.reject(new Error("JSZip can't accept a stream when loading a zip file.")):i.prepareContent("the loaded zip file",t,!0,s.optimizedBinaryString,s.base64).then(function(t){var e=new h(s);return e.load(t),e}).then(function(t){var e=[n.Promise.resolve(t)],r=t.files;if(s.checkCRC32)for(var i=0;i<r.length;i++)e.push(l(r[i]));return n.Promise.all(e)}).then(function(t){for(var e=t.shift(),r=e.files,i=0;i<r.length;i++){var n=r[i];a.file(n.fileNameStr,n.decompressed,{binary:!0,optimizedBinaryString:!0,date:n.date,dir:n.dir,comment:n.fileCommentStr.length?n.fileCommentStr:null,unixPermissions:n.unixPermissions,dosPermissions:n.dosPermissions,createFolders:s.createFolders})}return e.zipComment.length&&(a.comment=e.zipComment),a})}},{"./external":6,"./nodejsUtils":14,"./stream/Crc32Probe":25,"./utf8":31,"./utils":32,"./zipEntries":33}],12:[function(t,e,r){"use strict";var i=t("../utils"),n=t("../stream/GenericWorker");function s(t,e){n.call(this,"Nodejs stream input adapter for "+t),this._upstreamEnded=!1,this._bindStream(e)}i.inherits(s,n),s.prototype._bindStream=function(t){var e=this;(this._stream=t).pause(),t.on("data",function(t){e.push({data:t,meta:{percent:0}})}).on("error",function(t){e.isPaused?this.generatedError=t:e.error(t)}).on("end",function(){e.isPaused?e._upstreamEnded=!0:e.end()})},s.prototype.pause=function(){return!!n.prototype.pause.call(this)&&(this._stream.pause(),!0)},s.prototype.resume=function(){return!!n.prototype.resume.call(this)&&(this._upstreamEnded?this.end():this._stream.resume(),!0)},e.exports=s},{"../stream/GenericWorker":28,"../utils":32}],13:[function(t,e,r){"use strict";var n=t("readable-stream").Readable;function i(t,e,r){n.call(this,e),this._helper=t;var i=this;t.on("data",function(t,e){i.push(t)||i._helper.pause(),r&&r(e)}).on("error",function(t){i.emit("error",t)}).on("end",function(){i.push(null)})}t("../utils").inherits(i,n),i.prototype._read=function(){this._helper.resume()},e.exports=i},{"../utils":32,"readable-stream":16}],14:[function(t,e,r){"use strict";e.exports={isNode:"undefined"!=typeof Buffer,newBufferFrom:function(t,e){if(Buffer.from&&Buffer.from!==Uint8Array.from)return Buffer.from(t,e);if("number"==typeof t)throw new Error('The "data" argument must not be a number');return new Buffer(t,e)},allocBuffer:function(t){if(Buffer.alloc)return Buffer.alloc(t);var e=new Buffer(t);return e.fill(0),e},isBuffer:function(t){return Buffer.isBuffer(t)},isStream:function(t){return t&&"function"==typeof t.on&&"function"==typeof t.pause&&"function"==typeof t.resume}}},{}],15:[function(t,e,r){"use strict";function s(t,e,r){var i,n=u.getTypeOf(e),s=u.extend(r||{},f);s.date=s.date||new Date,null!==s.compression&&(s.compression=s.compression.toUpperCase()),"string"==typeof s.unixPermissions&&(s.unixPermissions=parseInt(s.unixPermissions,8)),s.unixPermissions&&16384&s.unixPermissions&&(s.dir=!0),s.dosPermissions&&16&s.dosPermissions&&(s.dir=!0),s.dir&&(t=g(t)),s.createFolders&&(i=_(t))&&b.call(this,i,!0);var a="string"===n&&!1===s.binary&&!1===s.base64;r&&void 0!==r.binary||(s.binary=!a),(e instanceof d&&0===e.uncompressedSize||s.dir||!e||0===e.length)&&(s.base64=!1,s.binary=!0,e="",s.compression="STORE",n="string");var o=null;o=e instanceof d||e instanceof l?e:p.isNode&&p.isStream(e)?new m(t,e):u.prepareContent(t,e,s.binary,s.optimizedBinaryString,s.base64);var h=new c(t,o,s);this.files[t]=h}var n=t("./utf8"),u=t("./utils"),l=t("./stream/GenericWorker"),a=t("./stream/StreamHelper"),f=t("./defaults"),d=t("./compressedObject"),c=t("./zipObject"),o=t("./generate"),p=t("./nodejsUtils"),m=t("./nodejs/NodejsStreamInputAdapter"),_=function(t){"/"===t.slice(-1)&&(t=t.substring(0,t.length-1));var e=t.lastIndexOf("/");return 0<e?t.substring(0,e):""},g=function(t){return"/"!==t.slice(-1)&&(t+="/"),t},b=function(t,e){return e=void 0!==e?e:f.createFolders,t=g(t),this.files[t]||s.call(this,t,null,{dir:!0,createFolders:e}),this.files[t]};function h(t){return"[object RegExp]"===Object.prototype.toString.call(t)}var i={load:function(){throw new Error("This method has been removed in JSZip 3.0, please check the upgrade guide.")},forEach:function(t){var e,r,i;for(e in this.files)i=this.files[e],(r=e.slice(this.root.length,e.length))&&e.slice(0,this.root.length)===this.root&&t(r,i)},filter:function(r){var i=[];return this.forEach(function(t,e){r(t,e)&&i.push(e)}),i},file:function(t,e,r){if(1!==arguments.length)return t=this.root+t,s.call(this,t,e,r),this;if(h(t)){var i=t;return this.filter(function(t,e){return!e.dir&&i.test(t)})}var n=this.files[this.root+t];return n&&!n.dir?n:null},folder:function(r){if(!r)return this;if(h(r))return this.filter(function(t,e){return e.dir&&r.test(t)});var t=this.root+r,e=b.call(this,t),i=this.clone();return i.root=e.name,i},remove:function(r){r=this.root+r;var t=this.files[r];if(t||("/"!==r.slice(-1)&&(r+="/"),t=this.files[r]),t&&!t.dir)delete this.files[r];else for(var e=this.filter(function(t,e){return e.name.slice(0,r.length)===r}),i=0;i<e.length;i++)delete this.files[e[i].name];return this},generate:function(t){throw new Error("This method has been removed in JSZip 3.0, please check the upgrade guide.")},generateInternalStream:function(t){var e,r={};try{if((r=u.extend(t||{},{streamFiles:!1,compression:"STORE",compressionOptions:null,type:"",platform:"DOS",comment:null,mimeType:"application/zip",encodeFileName:n.utf8encode})).type=r.type.toLowerCase(),r.compression=r.compression.toUpperCase(),"binarystring"===r.type&&(r.type="string"),!r.type)throw new Error("No output type specified.");u.checkSupport(r.type),"darwin"!==r.platform&&"freebsd"!==r.platform&&"linux"!==r.platform&&"sunos"!==r.platform||(r.platform="UNIX"),"win32"===r.platform&&(r.platform="DOS");var i=r.comment||this.comment||"";e=o.generateWorker(this,r,i)}catch(t){(e=new l("error")).error(t)}return new a(e,r.type||"string",r.mimeType)},generateAsync:function(t,e){return this.generateInternalStream(t).accumulate(e)},generateNodeStream:function(t,e){return(t=t||{}).type||(t.type="nodebuffer"),this.generateInternalStream(t).toNodejsStream(e)}};e.exports=i},{"./compressedObject":2,"./defaults":5,"./generate":9,"./nodejs/NodejsStreamInputAdapter":12,"./nodejsUtils":14,"./stream/GenericWorker":28,"./stream/StreamHelper":29,"./utf8":31,"./utils":32,"./zipObject":35}],16:[function(t,e,r){e.exports=t("stream")},{stream:void 0}],17:[function(t,e,r){"use strict";var i=t("./DataReader");function n(t){i.call(this,t);for(var e=0;e<this.data.length;e++)t[e]=255&t[e]}t("../utils").inherits(n,i),n.prototype.byteAt=function(t){return this.data[this.zero+t]},n.prototype.lastIndexOfSignature=function(t){for(var e=t.charCodeAt(0),r=t.charCodeAt(1),i=t.charCodeAt(2),n=t.charCodeAt(3),s=this.length-4;0<=s;--s)if(this.data[s]===e&&this.data[s+1]===r&&this.data[s+2]===i&&this.data[s+3]===n)return s-this.zero;return-1},n.prototype.readAndCheckSignature=function(t){var e=t.charCodeAt(0),r=t.charCodeAt(1),i=t.charCodeAt(2),n=t.charCodeAt(3),s=this.readData(4);return e===s[0]&&r===s[1]&&i===s[2]&&n===s[3]},n.prototype.readData=function(t){if(this.checkOffset(t),0===t)return[];var e=this.data.slice(this.zero+this.index,this.zero+this.index+t);return this.index+=t,e},e.exports=n},{"../utils":32,"./DataReader":18}],18:[function(t,e,r){"use strict";var i=t("../utils");function n(t){this.data=t,this.length=t.length,this.index=0,this.zero=0}n.prototype={checkOffset:function(t){this.checkIndex(this.index+t)},checkIndex:function(t){if(this.length<this.zero+t||t<0)throw new Error("End of data reached (data length = "+this.length+", asked index = "+t+"). Corrupted zip ?")},setIndex:function(t){this.checkIndex(t),this.index=t},skip:function(t){this.setIndex(this.index+t)},byteAt:function(t){},readInt:function(t){var e,r=0;for(this.checkOffset(t),e=this.index+t-1;e>=this.index;e--)r=(r<<8)+this.byteAt(e);return this.index+=t,r},readString:function(t){return i.transformTo("string",this.readData(t))},readData:function(t){},lastIndexOfSignature:function(t){},readAndCheckSignature:function(t){},readDate:function(){var t=this.readInt(4);return new Date(Date.UTC(1980+(t>>25&127),(t>>21&15)-1,t>>16&31,t>>11&31,t>>5&63,(31&t)<<1))}},e.exports=n},{"../utils":32}],19:[function(t,e,r){"use strict";var i=t("./Uint8ArrayReader");function n(t){i.call(this,t)}t("../utils").inherits(n,i),n.prototype.readData=function(t){this.checkOffset(t);var e=this.data.slice(this.zero+this.index,this.zero+this.index+t);return this.index+=t,e},e.exports=n},{"../utils":32,"./Uint8ArrayReader":21}],20:[function(t,e,r){"use strict";var i=t("./DataReader");function n(t){i.call(this,t)}t("../utils").inherits(n,i),n.prototype.byteAt=function(t){return this.data.charCodeAt(this.zero+t)},n.prototype.lastIndexOfSignature=function(t){return this.data.lastIndexOf(t)-this.zero},n.prototype.readAndCheckSignature=function(t){return t===this.readData(4)},n.prototype.readData=function(t){this.checkOffset(t);var e=this.data.slice(this.zero+this.index,this.zero+this.index+t);return this.index+=t,e},e.exports=n},{"../utils":32,"./DataReader":18}],21:[function(t,e,r){"use strict";var i=t("./ArrayReader");function n(t){i.call(this,t)}t("../utils").inherits(n,i),n.prototype.readData=function(t){if(this.checkOffset(t),0===t)return new Uint8Array(0);var e=this.data.subarray(this.zero+this.index,this.zero+this.index+t);return this.index+=t,e},e.exports=n},{"../utils":32,"./ArrayReader":17}],22:[function(t,e,r){"use strict";var i=t("../utils"),n=t("../support"),s=t("./ArrayReader"),a=t("./StringReader"),o=t("./NodeBufferReader"),h=t("./Uint8ArrayReader");e.exports=function(t){var e=i.getTypeOf(t);return i.checkSupport(e),"string"!==e||n.uint8array?"nodebuffer"===e?new o(t):n.uint8array?new h(i.transformTo("uint8array",t)):new s(i.transformTo("array",t)):new a(t)}},{"../support":30,"../utils":32,"./ArrayReader":17,"./NodeBufferReader":19,"./StringReader":20,"./Uint8ArrayReader":21}],23:[function(t,e,r){"use strict";r.LOCAL_FILE_HEADER="PK",r.CENTRAL_FILE_HEADER="PK",r.CENTRAL_DIRECTORY_END="PK",r.ZIP64_CENTRAL_DIRECTORY_LOCATOR="PK",r.ZIP64_CENTRAL_DIRECTORY_END="PK",r.DATA_DESCRIPTOR="PK\b"},{}],24:[function(t,e,r){"use strict";var i=t("./GenericWorker"),n=t("../utils");function s(t){i.call(this,"ConvertWorker to "+t),this.destType=t}n.inherits(s,i),s.prototype.processChunk=function(t){this.push({data:n.transformTo(this.destType,t.data),meta:t.meta})},e.exports=s},{"../utils":32,"./GenericWorker":28}],25:[function(t,e,r){"use strict";var i=t("./GenericWorker"),n=t("../crc32");function s(){i.call(this,"Crc32Probe"),this.withStreamInfo("crc32",0)}t("../utils").inherits(s,i),s.prototype.processChunk=function(t){this.streamInfo.crc32=n(t.data,this.streamInfo.crc32||0),this.push(t)},e.exports=s},{"../crc32":4,"../utils":32,"./GenericWorker":28}],26:[function(t,e,r){"use strict";var i=t("../utils"),n=t("./GenericWorker");function s(t){n.call(this,"DataLengthProbe for "+t),this.propName=t,this.withStreamInfo(t,0)}i.inherits(s,n),s.prototype.processChunk=function(t){if(t){var e=this.streamInfo[this.propName]||0;this.streamInfo[this.propName]=e+t.data.length}n.prototype.processChunk.call(this,t)},e.exports=s},{"../utils":32,"./GenericWorker":28}],27:[function(t,e,r){"use strict";var i=t("../utils"),n=t("./GenericWorker");function s(t){n.call(this,"DataWorker");var e=this;this.dataIsReady=!1,this.index=0,this.max=0,this.data=null,this.type="",this._tickScheduled=!1,t.then(function(t){e.dataIsReady=!0,e.data=t,e.max=t&&t.length||0,e.type=i.getTypeOf(t),e.isPaused||e._tickAndRepeat()},function(t){e.error(t)})}i.inherits(s,n),s.prototype.cleanUp=function(){n.prototype.cleanUp.call(this),this.data=null},s.prototype.resume=function(){return!!n.prototype.resume.call(this)&&(!this._tickScheduled&&this.dataIsReady&&(this._tickScheduled=!0,i.delay(this._tickAndRepeat,[],this)),!0)},s.prototype._tickAndRepeat=function(){this._tickScheduled=!1,this.isPaused||this.isFinished||(this._tick(),this.isFinished||(i.delay(this._tickAndRepeat,[],this),this._tickScheduled=!0))},s.prototype._tick=function(){if(this.isPaused||this.isFinished)return!1;var t=null,e=Math.min(this.max,this.index+16384);if(this.index>=this.max)return this.end();switch(this.type){case"string":t=this.data.substring(this.index,e);break;case"uint8array":t=this.data.subarray(this.index,e);break;case"array":case"nodebuffer":t=this.data.slice(this.index,e)}return this.index=e,this.push({data:t,meta:{percent:this.max?this.index/this.max*100:0}})},e.exports=s},{"../utils":32,"./GenericWorker":28}],28:[function(t,e,r){"use strict";function i(t){this.name=t||"default",this.streamInfo={},this.generatedError=null,this.extraStreamInfo={},this.isPaused=!0,this.isFinished=!1,this.isLocked=!1,this._listeners={data:[],end:[],error:[]},this.previous=null}i.prototype={push:function(t){this.emit("data",t)},end:function(){if(this.isFinished)return!1;this.flush();try{this.emit("end"),this.cleanUp(),this.isFinished=!0}catch(t){this.emit("error",t)}return!0},error:function(t){return!this.isFinished&&(this.isPaused?this.generatedError=t:(this.isFinished=!0,this.emit("error",t),this.previous&&this.previous.error(t),this.cleanUp()),!0)},on:function(t,e){return this._listeners[t].push(e),this},cleanUp:function(){this.streamInfo=this.generatedError=this.extraStreamInfo=null,this._listeners=[]},emit:function(t,e){if(this._listeners[t])for(var r=0;r<this._listeners[t].length;r++)this._listeners[t][r].call(this,e)},pipe:function(t){return t.registerPrevious(this)},registerPrevious:function(t){if(this.isLocked)throw new Error("The stream '"+this+"' has already been used.");this.streamInfo=t.streamInfo,this.mergeStreamInfo(),this.previous=t;var e=this;return t.on("data",function(t){e.processChunk(t)}),t.on("end",function(){e.end()}),t.on("error",function(t){e.error(t)}),this},pause:function(){return!this.isPaused&&!this.isFinished&&(this.isPaused=!0,this.previous&&this.previous.pause(),!0)},resume:function(){if(!this.isPaused||this.isFinished)return!1;var t=this.isPaused=!1;return this.generatedError&&(this.error(this.generatedError),t=!0),this.previous&&this.previous.resume(),!t},flush:function(){},processChunk:function(t){this.push(t)},withStreamInfo:function(t,e){return this.extraStreamInfo[t]=e,this.mergeStreamInfo(),this},mergeStreamInfo:function(){for(var t in this.extraStreamInfo)this.extraStreamInfo.hasOwnProperty(t)&&(this.streamInfo[t]=this.extraStreamInfo[t])},lock:function(){if(this.isLocked)throw new Error("The stream '"+this+"' has already been used.");this.isLocked=!0,this.previous&&this.previous.lock()},toString:function(){var t="Worker "+this.name;return this.previous?this.previous+" -> "+t:t}},e.exports=i},{}],29:[function(t,e,r){"use strict";var h=t("../utils"),n=t("./ConvertWorker"),s=t("./GenericWorker"),u=t("../base64"),i=t("../support"),a=t("../external"),o=null;if(i.nodestream)try{o=t("../nodejs/NodejsStreamOutputAdapter")}catch(t){}function l(t,o){return new a.Promise(function(e,r){var i=[],n=t._internalType,s=t._outputType,a=t._mimeType;t.on("data",function(t,e){i.push(t),o&&o(e)}).on("error",function(t){i=[],r(t)}).on("end",function(){try{var t=function(t,e,r){switch(t){case"blob":return h.newBlob(h.transformTo("arraybuffer",e),r);case"base64":return u.encode(e);default:return h.transformTo(t,e)}}(s,function(t,e){var r,i=0,n=null,s=0;for(r=0;r<e.length;r++)s+=e[r].length;switch(t){case"string":return e.join("");case"array":return Array.prototype.concat.apply([],e);case"uint8array":for(n=new Uint8Array(s),r=0;r<e.length;r++)n.set(e[r],i),i+=e[r].length;return n;case"nodebuffer":return Buffer.concat(e);default:throw new Error("concat : unsupported type '"+t+"'")}}(n,i),a);e(t)}catch(t){r(t)}i=[]}).resume()})}function f(t,e,r){var i=e;switch(e){case"blob":case"arraybuffer":i="uint8array";break;case"base64":i="string"}try{this._internalType=i,this._outputType=e,this._mimeType=r,h.checkSupport(i),this._worker=t.pipe(new n(i)),t.lock()}catch(t){this._worker=new s("error"),this._worker.error(t)}}f.prototype={accumulate:function(t){return l(this,t)},on:function(t,e){var r=this;return"data"===t?this._worker.on(t,function(t){e.call(r,t.data,t.meta)}):this._worker.on(t,function(){h.delay(e,arguments,r)}),this},resume:function(){return h.delay(this._worker.resume,[],this._worker),this},pause:function(){return this._worker.pause(),this},toNodejsStream:function(t){if(h.checkSupport("nodestream"),"nodebuffer"!==this._outputType)throw new Error(this._outputType+" is not supported by this method");return new o(this,{objectMode:"nodebuffer"!==this._outputType},t)}},e.exports=f},{"../base64":1,"../external":6,"../nodejs/NodejsStreamOutputAdapter":13,"../support":30,"../utils":32,"./ConvertWorker":24,"./GenericWorker":28}],30:[function(t,e,r){"use strict";if(r.base64=!0,r.array=!0,r.string=!0,r.arraybuffer="undefined"!=typeof ArrayBuffer&&"undefined"!=typeof Uint8Array,r.nodebuffer="undefined"!=typeof Buffer,r.uint8array="undefined"!=typeof Uint8Array,"undefined"==typeof ArrayBuffer)r.blob=!1;else{var i=new ArrayBuffer(0);try{r.blob=0===new Blob([i],{type:"application/zip"}).size}catch(t){try{var n=new(self.BlobBuilder||self.WebKitBlobBuilder||self.MozBlobBuilder||self.MSBlobBuilder);n.append(i),r.blob=0===n.getBlob("application/zip").size}catch(t){r.blob=!1}}}try{r.nodestream=!!t("readable-stream").Readable}catch(t){r.nodestream=!1}},{"readable-stream":16}],31:[function(t,e,s){"use strict";for(var o=t("./utils"),h=t("./support"),r=t("./nodejsUtils"),i=t("./stream/GenericWorker"),u=new Array(256),n=0;n<256;n++)u[n]=252<=n?6:248<=n?5:240<=n?4:224<=n?3:192<=n?2:1;u[254]=u[254]=1;function a(){i.call(this,"utf-8 decode"),this.leftOver=null}function l(){i.call(this,"utf-8 encode")}s.utf8encode=function(t){return h.nodebuffer?r.newBufferFrom(t,"utf-8"):function(t){var e,r,i,n,s,a=t.length,o=0;for(n=0;n<a;n++)55296==(64512&(r=t.charCodeAt(n)))&&n+1<a&&56320==(64512&(i=t.charCodeAt(n+1)))&&(r=65536+(r-55296<<10)+(i-56320),n++),o+=r<128?1:r<2048?2:r<65536?3:4;for(e=h.uint8array?new Uint8Array(o):new Array(o),n=s=0;s<o;n++)55296==(64512&(r=t.charCodeAt(n)))&&n+1<a&&56320==(64512&(i=t.charCodeAt(n+1)))&&(r=65536+(r-55296<<10)+(i-56320),n++),r<128?e[s++]=r:(r<2048?e[s++]=192|r>>>6:(r<65536?e[s++]=224|r>>>12:(e[s++]=240|r>>>18,e[s++]=128|r>>>12&63),e[s++]=128|r>>>6&63),e[s++]=128|63&r);return e}(t)},s.utf8decode=function(t){return h.nodebuffer?o.transformTo("nodebuffer",t).toString("utf-8"):function(t){var e,r,i,n,s=t.length,a=new Array(2*s);for(e=r=0;e<s;)if((i=t[e++])<128)a[r++]=i;else if(4<(n=u[i]))a[r++]=65533,e+=n-1;else{for(i&=2===n?31:3===n?15:7;1<n&&e<s;)i=i<<6|63&t[e++],n--;1<n?a[r++]=65533:i<65536?a[r++]=i:(i-=65536,a[r++]=55296|i>>10&1023,a[r++]=56320|1023&i)}return a.length!==r&&(a.subarray?a=a.subarray(0,r):a.length=r),o.applyFromCharCode(a)}(t=o.transformTo(h.uint8array?"uint8array":"array",t))},o.inherits(a,i),a.prototype.processChunk=function(t){var e=o.transformTo(h.uint8array?"uint8array":"array",t.data);if(this.leftOver&&this.leftOver.length){if(h.uint8array){var r=e;(e=new Uint8Array(r.length+this.leftOver.length)).set(this.leftOver,0),e.set(r,this.leftOver.length)}else e=this.leftOver.concat(e);this.leftOver=null}var i=function(t,e){var r;for((e=e||t.length)>t.length&&(e=t.length),r=e-1;0<=r&&128==(192&t[r]);)r--;return r<0?e:0===r?e:r+u[t[r]]>e?r:e}(e),n=e;i!==e.length&&(h.uint8array?(n=e.subarray(0,i),this.leftOver=e.subarray(i,e.length)):(n=e.slice(0,i),this.leftOver=e.slice(i,e.length))),this.push({data:s.utf8decode(n),meta:t.meta})},a.prototype.flush=function(){this.leftOver&&this.leftOver.length&&(this.push({data:s.utf8decode(this.leftOver),meta:{}}),this.leftOver=null)},s.Utf8DecodeWorker=a,o.inherits(l,i),l.prototype.processChunk=function(t){this.push({data:s.utf8encode(t.data),meta:t.meta})},s.Utf8EncodeWorker=l},{"./nodejsUtils":14,"./stream/GenericWorker":28,"./support":30,"./utils":32}],32:[function(t,e,a){"use strict";var o=t("./support"),h=t("./base64"),r=t("./nodejsUtils"),i=t("set-immediate-shim"),u=t("./external");function n(t){return t}function l(t,e){for(var r=0;r<t.length;++r)e[r]=255&t.charCodeAt(r);return e}a.newBlob=function(e,r){a.checkSupport("blob");try{return new Blob([e],{type:r})}catch(t){try{var i=new(self.BlobBuilder||self.WebKitBlobBuilder||self.MozBlobBuilder||self.MSBlobBuilder);return i.append(e),i.getBlob(r)}catch(t){throw new Error("Bug : can't construct the Blob.")}}};var s={stringifyByChunk:function(t,e,r){var i=[],n=0,s=t.length;if(s<=r)return String.fromCharCode.apply(null,t);for(;n<s;)"array"===e||"nodebuffer"===e?i.push(String.fromCharCode.apply(null,t.slice(n,Math.min(n+r,s)))):i.push(String.fromCharCode.apply(null,t.subarray(n,Math.min(n+r,s)))),n+=r;return i.join("")},stringifyByChar:function(t){for(var e="",r=0;r<t.length;r++)e+=String.fromCharCode(t[r]);return e},applyCanBeUsed:{uint8array:function(){try{return o.uint8array&&1===String.fromCharCode.apply(null,new Uint8Array(1)).length}catch(t){return!1}}(),nodebuffer:function(){try{return o.nodebuffer&&1===String.fromCharCode.apply(null,r.allocBuffer(1)).length}catch(t){return!1}}()}};function f(t){var e=65536,r=a.getTypeOf(t),i=!0;if("uint8array"===r?i=s.applyCanBeUsed.uint8array:"nodebuffer"===r&&(i=s.applyCanBeUsed.nodebuffer),i)for(;1<e;)try{return s.stringifyByChunk(t,r,e)}catch(t){e=Math.floor(e/2)}return s.stringifyByChar(t)}function d(t,e){for(var r=0;r<t.length;r++)e[r]=t[r];return e}a.applyFromCharCode=f;var c={};c.string={string:n,array:function(t){return l(t,new Array(t.length))},arraybuffer:function(t){return c.string.uint8array(t).buffer},uint8array:function(t){return l(t,new Uint8Array(t.length))},nodebuffer:function(t){return l(t,r.allocBuffer(t.length))}},c.array={string:f,array:n,arraybuffer:function(t){return new Uint8Array(t).buffer},uint8array:function(t){return new Uint8Array(t)},nodebuffer:function(t){return r.newBufferFrom(t)}},c.arraybuffer={string:function(t){return f(new Uint8Array(t))},array:function(t){return d(new Uint8Array(t),new Array(t.byteLength))},arraybuffer:n,uint8array:function(t){return new Uint8Array(t)},nodebuffer:function(t){return r.newBufferFrom(new Uint8Array(t))}},c.uint8array={string:f,array:function(t){return d(t,new Array(t.length))},arraybuffer:function(t){return t.buffer},uint8array:n,nodebuffer:function(t){return r.newBufferFrom(t)}},c.nodebuffer={string:f,array:function(t){return d(t,new Array(t.length))},arraybuffer:function(t){return c.nodebuffer.uint8array(t).buffer},uint8array:function(t){return d(t,new Uint8Array(t.length))},nodebuffer:n},a.transformTo=function(t,e){if(e=e||"",!t)return e;a.checkSupport(t);var r=a.getTypeOf(e);return c[r][t](e)},a.getTypeOf=function(t){return"string"==typeof t?"string":"[object Array]"===Object.prototype.toString.call(t)?"array":o.nodebuffer&&r.isBuffer(t)?"nodebuffer":o.uint8array&&t instanceof Uint8Array?"uint8array":o.arraybuffer&&t instanceof ArrayBuffer?"arraybuffer":void 0},a.checkSupport=function(t){if(!o[t.toLowerCase()])throw new Error(t+" is not supported by this platform")},a.MAX_VALUE_16BITS=65535,a.MAX_VALUE_32BITS=-1,a.pretty=function(t){var e,r,i="";for(r=0;r<(t||"").length;r++)i+="\\x"+((e=t.charCodeAt(r))<16?"0":"")+e.toString(16).toUpperCase();return i},a.delay=function(t,e,r){i(function(){t.apply(r||null,e||[])})},a.inherits=function(t,e){function r(){}r.prototype=e.prototype,t.prototype=new r},a.extend=function(){var t,e,r={};for(t=0;t<arguments.length;t++)for(e in arguments[t])arguments[t].hasOwnProperty(e)&&void 0===r[e]&&(r[e]=arguments[t][e]);return r},a.prepareContent=function(r,t,i,n,s){return u.Promise.resolve(t).then(function(i){return o.blob&&(i instanceof Blob||-1!==["[object File]","[object Blob]"].indexOf(Object.prototype.toString.call(i)))&&"undefined"!=typeof FileReader?new u.Promise(function(e,r){var t=new FileReader;t.onload=function(t){e(t.target.result)},t.onerror=function(t){r(t.target.error)},t.readAsArrayBuffer(i)}):i}).then(function(t){var e=a.getTypeOf(t);return e?("arraybuffer"===e?t=a.transformTo("uint8array",t):"string"===e&&(s?t=h.decode(t):i&&!0!==n&&(t=function(t){return l(t,o.uint8array?new Uint8Array(t.length):new Array(t.length))}(t))),t):u.Promise.reject(new Error("Can't read the data of '"+r+"'. Is it in a supported JavaScript type (String, Blob, ArrayBuffer, etc) ?"))})}},{"./base64":1,"./external":6,"./nodejsUtils":14,"./support":30,"set-immediate-shim":54}],33:[function(t,e,r){"use strict";var i=t("./reader/readerFor"),n=t("./utils"),s=t("./signature"),a=t("./zipEntry"),o=(t("./utf8"),t("./support"));function h(t){this.files=[],this.loadOptions=t}h.prototype={checkSignature:function(t){if(!this.reader.readAndCheckSignature(t)){this.reader.index-=4;var e=this.reader.readString(4);throw new Error("Corrupted zip or bug: unexpected signature ("+n.pretty(e)+", expected "+n.pretty(t)+")")}},isSignature:function(t,e){var r=this.reader.index;this.reader.setIndex(t);var i=this.reader.readString(4)===e;return this.reader.setIndex(r),i},readBlockEndOfCentral:function(){this.diskNumber=this.reader.readInt(2),this.diskWithCentralDirStart=this.reader.readInt(2),this.centralDirRecordsOnThisDisk=this.reader.readInt(2),this.centralDirRecords=this.reader.readInt(2),this.centralDirSize=this.reader.readInt(4),this.centralDirOffset=this.reader.readInt(4),this.zipCommentLength=this.reader.readInt(2);var t=this.reader.readData(this.zipCommentLength),e=o.uint8array?"uint8array":"array",r=n.transformTo(e,t);this.zipComment=this.loadOptions.decodeFileName(r)},readBlockZip64EndOfCentral:function(){this.zip64EndOfCentralSize=this.reader.readInt(8),this.reader.skip(4),this.diskNumber=this.reader.readInt(4),this.diskWithCentralDirStart=this.reader.readInt(4),this.centralDirRecordsOnThisDisk=this.reader.readInt(8),this.centralDirRecords=this.reader.readInt(8),this.centralDirSize=this.reader.readInt(8),this.centralDirOffset=this.reader.readInt(8),this.zip64ExtensibleData={};for(var t,e,r,i=this.zip64EndOfCentralSize-44;0<i;)t=this.reader.readInt(2),e=this.reader.readInt(4),r=this.reader.readData(e),this.zip64ExtensibleData[t]={id:t,length:e,value:r}},readBlockZip64EndOfCentralLocator:function(){if(this.diskWithZip64CentralDirStart=this.reader.readInt(4),this.relativeOffsetEndOfZip64CentralDir=this.reader.readInt(8),this.disksCount=this.reader.readInt(4),1<this.disksCount)throw new Error("Multi-volumes zip are not supported")},readLocalFiles:function(){var t,e;for(t=0;t<this.files.length;t++)e=this.files[t],this.reader.setIndex(e.localHeaderOffset),this.checkSignature(s.LOCAL_FILE_HEADER),e.readLocalPart(this.reader),e.handleUTF8(),e.processAttributes()},readCentralDir:function(){var t;for(this.reader.setIndex(this.centralDirOffset);this.reader.readAndCheckSignature(s.CENTRAL_FILE_HEADER);)(t=new a({zip64:this.zip64},this.loadOptions)).readCentralPart(this.reader),this.files.push(t);if(this.centralDirRecords!==this.files.length&&0!==this.centralDirRecords&&0===this.files.length)throw new Error("Corrupted zip or bug: expected "+this.centralDirRecords+" records in central dir, got "+this.files.length)},readEndOfCentral:function(){var t=this.reader.lastIndexOfSignature(s.CENTRAL_DIRECTORY_END);if(t<0)throw!this.isSignature(0,s.LOCAL_FILE_HEADER)?new Error("Can't find end of central directory : is this a zip file ? If it is, see https://stuk.github.io/jszip/documentation/howto/read_zip.html"):new Error("Corrupted zip: can't find end of central directory");this.reader.setIndex(t);var e=t;if(this.checkSignature(s.CENTRAL_DIRECTORY_END),this.readBlockEndOfCentral(),this.diskNumber===n.MAX_VALUE_16BITS||this.diskWithCentralDirStart===n.MAX_VALUE_16BITS||this.centralDirRecordsOnThisDisk===n.MAX_VALUE_16BITS||this.centralDirRecords===n.MAX_VALUE_16BITS||this.centralDirSize===n.MAX_VALUE_32BITS||this.centralDirOffset===n.MAX_VALUE_32BITS){if(this.zip64=!0,(t=this.reader.lastIndexOfSignature(s.ZIP64_CENTRAL_DIRECTORY_LOCATOR))<0)throw new Error("Corrupted zip: can't find the ZIP64 end of central directory locator");if(this.reader.setIndex(t),this.checkSignature(s.ZIP64_CENTRAL_DIRECTORY_LOCATOR),this.readBlockZip64EndOfCentralLocator(),!this.isSignature(this.relativeOffsetEndOfZip64CentralDir,s.ZIP64_CENTRAL_DIRECTORY_END)&&(this.relativeOffsetEndOfZip64CentralDir=this.reader.lastIndexOfSignature(s.ZIP64_CENTRAL_DIRECTORY_END),this.relativeOffsetEndOfZip64CentralDir<0))throw new Error("Corrupted zip: can't find the ZIP64 end of central directory");this.reader.setIndex(this.relativeOffsetEndOfZip64CentralDir),this.checkSignature(s.ZIP64_CENTRAL_DIRECTORY_END),this.readBlockZip64EndOfCentral()}var r=this.centralDirOffset+this.centralDirSize;this.zip64&&(r+=20,r+=12+this.zip64EndOfCentralSize);var i=e-r;if(0<i)this.isSignature(e,s.CENTRAL_FILE_HEADER)||(this.reader.zero=i);else if(i<0)throw new Error("Corrupted zip: missing "+Math.abs(i)+" bytes.")},prepareReader:function(t){this.reader=i(t)},load:function(t){this.prepareReader(t),this.readEndOfCentral(),this.readCentralDir(),this.readLocalFiles()}},e.exports=h},{"./reader/readerFor":22,"./signature":23,"./support":30,"./utf8":31,"./utils":32,"./zipEntry":34}],34:[function(t,e,r){"use strict";var i=t("./reader/readerFor"),s=t("./utils"),n=t("./compressedObject"),a=t("./crc32"),o=t("./utf8"),h=t("./compressions"),u=t("./support");function l(t,e){this.options=t,this.loadOptions=e}l.prototype={isEncrypted:function(){return 1==(1&this.bitFlag)},useUTF8:function(){return 2048==(2048&this.bitFlag)},readLocalPart:function(t){var e,r;if(t.skip(22),this.fileNameLength=t.readInt(2),r=t.readInt(2),this.fileName=t.readData(this.fileNameLength),t.skip(r),-1===this.compressedSize||-1===this.uncompressedSize)throw new Error("Bug or corrupted zip : didn't get enough information from the central directory (compressedSize === -1 || uncompressedSize === -1)");if(null===(e=function(t){for(var e in h)if(h.hasOwnProperty(e)&&h[e].magic===t)return h[e];return null}(this.compressionMethod)))throw new Error("Corrupted zip : compression "+s.pretty(this.compressionMethod)+" unknown (inner file : "+s.transformTo("string",this.fileName)+")");this.decompressed=new n(this.compressedSize,this.uncompressedSize,this.crc32,e,t.readData(this.compressedSize))},readCentralPart:function(t){this.versionMadeBy=t.readInt(2),t.skip(2),this.bitFlag=t.readInt(2),this.compressionMethod=t.readString(2),this.date=t.readDate(),this.crc32=t.readInt(4),this.compressedSize=t.readInt(4),this.uncompressedSize=t.readInt(4);var e=t.readInt(2);if(this.extraFieldsLength=t.readInt(2),this.fileCommentLength=t.readInt(2),this.diskNumberStart=t.readInt(2),this.internalFileAttributes=t.readInt(2),this.externalFileAttributes=t.readInt(4),this.localHeaderOffset=t.readInt(4),this.isEncrypted())throw new Error("Encrypted zip are not supported");t.skip(e),this.readExtraFields(t),this.parseZIP64ExtraField(t),this.fileComment=t.readData(this.fileCommentLength)},processAttributes:function(){this.unixPermissions=null,this.dosPermissions=null;var t=this.versionMadeBy>>8;this.dir=!!(16&this.externalFileAttributes),0==t&&(this.dosPermissions=63&this.externalFileAttributes),3==t&&(this.unixPermissions=this.externalFileAttributes>>16&65535),this.dir||"/"!==this.fileNameStr.slice(-1)||(this.dir=!0)},parseZIP64ExtraField:function(t){if(this.extraFields[1]){var e=i(this.extraFields[1].value);this.uncompressedSize===s.MAX_VALUE_32BITS&&(this.uncompressedSize=e.readInt(8)),this.compressedSize===s.MAX_VALUE_32BITS&&(this.compressedSize=e.readInt(8)),this.localHeaderOffset===s.MAX_VALUE_32BITS&&(this.localHeaderOffset=e.readInt(8)),this.diskNumberStart===s.MAX_VALUE_32BITS&&(this.diskNumberStart=e.readInt(4))}},readExtraFields:function(t){var e,r,i,n=t.index+this.extraFieldsLength;for(this.extraFields||(this.extraFields={});t.index+4<n;)e=t.readInt(2),r=t.readInt(2),i=t.readData(r),this.extraFields[e]={id:e,length:r,value:i};t.setIndex(n)},handleUTF8:function(){var t=u.uint8array?"uint8array":"array";if(this.useUTF8())this.fileNameStr=o.utf8decode(this.fileName),this.fileCommentStr=o.utf8decode(this.fileComment);else{var e=this.findExtraFieldUnicodePath();if(null!==e)this.fileNameStr=e;else{var r=s.transformTo(t,this.fileName);this.fileNameStr=this.loadOptions.decodeFileName(r)}var i=this.findExtraFieldUnicodeComment();if(null!==i)this.fileCommentStr=i;else{var n=s.transformTo(t,this.fileComment);this.fileCommentStr=this.loadOptions.decodeFileName(n)}}},findExtraFieldUnicodePath:function(){var t=this.extraFields[28789];if(t){var e=i(t.value);return 1!==e.readInt(1)?null:a(this.fileName)!==e.readInt(4)?null:o.utf8decode(e.readData(t.length-5))}return null},findExtraFieldUnicodeComment:function(){var t=this.extraFields[25461];if(t){var e=i(t.value);return 1!==e.readInt(1)?null:a(this.fileComment)!==e.readInt(4)?null:o.utf8decode(e.readData(t.length-5))}return null}},e.exports=l},{"./compressedObject":2,"./compressions":3,"./crc32":4,"./reader/readerFor":22,"./support":30,"./utf8":31,"./utils":32}],35:[function(t,e,r){"use strict";function i(t,e,r){this.name=t,this.dir=r.dir,this.date=r.date,this.comment=r.comment,this.unixPermissions=r.unixPermissions,this.dosPermissions=r.dosPermissions,this._data=e,this._dataBinary=r.binary,this.options={compression:r.compression,compressionOptions:r.compressionOptions}}var s=t("./stream/StreamHelper"),n=t("./stream/DataWorker"),a=t("./utf8"),o=t("./compressedObject"),h=t("./stream/GenericWorker");i.prototype={internalStream:function(t){var e=null,r="string";try{if(!t)throw new Error("No output type specified.");var i="string"===(r=t.toLowerCase())||"text"===r;"binarystring"!==r&&"text"!==r||(r="string"),e=this._decompressWorker();var n=!this._dataBinary;n&&!i&&(e=e.pipe(new a.Utf8EncodeWorker)),!n&&i&&(e=e.pipe(new a.Utf8DecodeWorker))}catch(t){(e=new h("error")).error(t)}return new s(e,r,"")},async:function(t,e){return this.internalStream(t).accumulate(e)},nodeStream:function(t,e){return this.internalStream(t||"nodebuffer").toNodejsStream(e)},_compressWorker:function(t,e){if(this._data instanceof o&&this._data.compression.magic===t.magic)return this._data.getCompressedWorker();var r=this._decompressWorker();return this._dataBinary||(r=r.pipe(new a.Utf8EncodeWorker)),o.createWorkerFrom(r,t,e)},_decompressWorker:function(){return this._data instanceof o?this._data.getContentWorker():this._data instanceof h?this._data:new n(this._data)}};for(var u=["asText","asBinary","asNodeBuffer","asUint8Array","asArrayBuffer"],l=function(){throw new Error("This method has been removed in JSZip 3.0, please check the upgrade guide.")},f=0;f<u.length;f++)i.prototype[u[f]]=l;e.exports=i},{"./compressedObject":2,"./stream/DataWorker":27,"./stream/GenericWorker":28,"./stream/StreamHelper":29,"./utf8":31}],36:[function(t,l,e){(function(e){"use strict";var r,i,t=e.MutationObserver||e.WebKitMutationObserver;if(t){var n=0,s=new t(u),a=e.document.createTextNode("");s.observe(a,{characterData:!0}),r=function(){a.data=n=++n%2}}else if(e.setImmediate||void 0===e.MessageChannel)r="document"in e&&"onreadystatechange"in e.document.createElement("script")?function(){var t=e.document.createElement("script");t.onreadystatechange=function(){u(),t.onreadystatechange=null,t.parentNode.removeChild(t),t=null},e.document.documentElement.appendChild(t)}:function(){setTimeout(u,0)};else{var o=new e.MessageChannel;o.port1.onmessage=u,r=function(){o.port2.postMessage(0)}}var h=[];function u(){var t,e;i=!0;for(var r=h.length;r;){for(e=h,h=[],t=-1;++t<r;)e[t]();r=h.length}i=!1}l.exports=function(t){1!==h.push(t)||i||r()}}).call(this,"undefined"!=typeof global?global:"undefined"!=typeof self?self:"undefined"!=typeof window?window:{})},{}],37:[function(t,e,r){"use strict";var n=t("immediate");function u(){}var l={},s=["REJECTED"],a=["FULFILLED"],i=["PENDING"];function o(t){if("function"!=typeof t)throw new TypeError("resolver must be a function");this.state=i,this.queue=[],this.outcome=void 0,t!==u&&c(this,t)}function h(t,e,r){this.promise=t,"function"==typeof e&&(this.onFulfilled=e,this.callFulfilled=this.otherCallFulfilled),"function"==typeof r&&(this.onRejected=r,this.callRejected=this.otherCallRejected)}function f(e,r,i){n(function(){var t;try{t=r(i)}catch(t){return l.reject(e,t)}t===e?l.reject(e,new TypeError("Cannot resolve promise with itself")):l.resolve(e,t)})}function d(t){var e=t&&t.then;if(t&&("object"==typeof t||"function"==typeof t)&&"function"==typeof e)return function(){e.apply(t,arguments)}}function c(e,t){var r=!1;function i(t){r||(r=!0,l.reject(e,t))}function n(t){r||(r=!0,l.resolve(e,t))}var s=p(function(){t(n,i)});"error"===s.status&&i(s.value)}function p(t,e){var r={};try{r.value=t(e),r.status="success"}catch(t){r.status="error",r.value=t}return r}(e.exports=o).prototype.finally=function(e){if("function"!=typeof e)return this;var r=this.constructor;return this.then(function(t){return r.resolve(e()).then(function(){return t})},function(t){return r.resolve(e()).then(function(){throw t})})},o.prototype.catch=function(t){return this.then(null,t)},o.prototype.then=function(t,e){if("function"!=typeof t&&this.state===a||"function"!=typeof e&&this.state===s)return this;var r=new this.constructor(u);this.state!==i?f(r,this.state===a?t:e,this.outcome):this.queue.push(new h(r,t,e));return r},h.prototype.callFulfilled=function(t){l.resolve(this.promise,t)},h.prototype.otherCallFulfilled=function(t){f(this.promise,this.onFulfilled,t)},h.prototype.callRejected=function(t){l.reject(this.promise,t)},h.prototype.otherCallRejected=function(t){f(this.promise,this.onRejected,t)},l.resolve=function(t,e){var r=p(d,e);if("error"===r.status)return l.reject(t,r.value);var i=r.value;if(i)c(t,i);else{t.state=a,t.outcome=e;for(var n=-1,s=t.queue.length;++n<s;)t.queue[n].callFulfilled(e)}return t},l.reject=function(t,e){t.state=s,t.outcome=e;for(var r=-1,i=t.queue.length;++r<i;)t.queue[r].callRejected(e);return t},o.resolve=function(t){if(t instanceof this)return t;return l.resolve(new this(u),t)},o.reject=function(t){var e=new this(u);return l.reject(e,t)},o.all=function(t){var r=this;if("[object Array]"!==Object.prototype.toString.call(t))return this.reject(new TypeError("must be an array"));var i=t.length,n=!1;if(!i)return this.resolve([]);var s=new Array(i),a=0,e=-1,o=new this(u);for(;++e<i;)h(t[e],e);return o;function h(t,e){r.resolve(t).then(function(t){s[e]=t,++a!==i||n||(n=!0,l.resolve(o,s))},function(t){n||(n=!0,l.reject(o,t))})}},o.race=function(t){var e=this;if("[object Array]"!==Object.prototype.toString.call(t))return this.reject(new TypeError("must be an array"));var r=t.length,i=!1;if(!r)return this.resolve([]);var n=-1,s=new this(u);for(;++n<r;)a=t[n],e.resolve(a).then(function(t){i||(i=!0,l.resolve(s,t))},function(t){i||(i=!0,l.reject(s,t))});var a;return s}},{immediate:36}],38:[function(t,e,r){"use strict";var i={};(0,t("./lib/utils/common").assign)(i,t("./lib/deflate"),t("./lib/inflate"),t("./lib/zlib/constants")),e.exports=i},{"./lib/deflate":39,"./lib/inflate":40,"./lib/utils/common":41,"./lib/zlib/constants":44}],39:[function(t,e,r){"use strict";var a=t("./zlib/deflate"),o=t("./utils/common"),h=t("./utils/strings"),n=t("./zlib/messages"),s=t("./zlib/zstream"),u=Object.prototype.toString,l=0,f=-1,d=0,c=8;function p(t){if(!(this instanceof p))return new p(t);this.options=o.assign({level:f,method:c,chunkSize:16384,windowBits:15,memLevel:8,strategy:d,to:""},t||{});var e=this.options;e.raw&&0<e.windowBits?e.windowBits=-e.windowBits:e.gzip&&0<e.windowBits&&e.windowBits<16&&(e.windowBits+=16),this.err=0,this.msg="",this.ended=!1,this.chunks=[],this.strm=new s,this.strm.avail_out=0;var r=a.deflateInit2(this.strm,e.level,e.method,e.windowBits,e.memLevel,e.strategy);if(r!==l)throw new Error(n[r]);if(e.header&&a.deflateSetHeader(this.strm,e.header),e.dictionary){var i;if(i="string"==typeof e.dictionary?h.string2buf(e.dictionary):"[object ArrayBuffer]"===u.call(e.dictionary)?new Uint8Array(e.dictionary):e.dictionary,(r=a.deflateSetDictionary(this.strm,i))!==l)throw new Error(n[r]);this._dict_set=!0}}function i(t,e){var r=new p(e);if(r.push(t,!0),r.err)throw r.msg||n[r.err];return r.result}p.prototype.push=function(t,e){var r,i,n=this.strm,s=this.options.chunkSize;if(this.ended)return!1;i=e===~~e?e:!0===e?4:0,"string"==typeof t?n.input=h.string2buf(t):"[object ArrayBuffer]"===u.call(t)?n.input=new Uint8Array(t):n.input=t,n.next_in=0,n.avail_in=n.input.length;do{if(0===n.avail_out&&(n.output=new o.Buf8(s),n.next_out=0,n.avail_out=s),1!==(r=a.deflate(n,i))&&r!==l)return this.onEnd(r),!(this.ended=!0);0!==n.avail_out&&(0!==n.avail_in||4!==i&&2!==i)||("string"===this.options.to?this.onData(h.buf2binstring(o.shrinkBuf(n.output,n.next_out))):this.onData(o.shrinkBuf(n.output,n.next_out)))}while((0<n.avail_in||0===n.avail_out)&&1!==r);return 4===i?(r=a.deflateEnd(this.strm),this.onEnd(r),this.ended=!0,r===l):2!==i||(this.onEnd(l),!(n.avail_out=0))},p.prototype.onData=function(t){this.chunks.push(t)},p.prototype.onEnd=function(t){t===l&&("string"===this.options.to?this.result=this.chunks.join(""):this.result=o.flattenChunks(this.chunks)),this.chunks=[],this.err=t,this.msg=this.strm.msg},r.Deflate=p,r.deflate=i,r.deflateRaw=function(t,e){return(e=e||{}).raw=!0,i(t,e)},r.gzip=function(t,e){return(e=e||{}).gzip=!0,i(t,e)}},{"./utils/common":41,"./utils/strings":42,"./zlib/deflate":46,"./zlib/messages":51,"./zlib/zstream":53}],40:[function(t,e,r){"use strict";var d=t("./zlib/inflate"),c=t("./utils/common"),p=t("./utils/strings"),m=t("./zlib/constants"),i=t("./zlib/messages"),n=t("./zlib/zstream"),s=t("./zlib/gzheader"),_=Object.prototype.toString;function a(t){if(!(this instanceof a))return new a(t);this.options=c.assign({chunkSize:16384,windowBits:0,to:""},t||{});var e=this.options;e.raw&&0<=e.windowBits&&e.windowBits<16&&(e.windowBits=-e.windowBits,0===e.windowBits&&(e.windowBits=-15)),!(0<=e.windowBits&&e.windowBits<16)||t&&t.windowBits||(e.windowBits+=32),15<e.windowBits&&e.windowBits<48&&0==(15&e.windowBits)&&(e.windowBits|=15),this.err=0,this.msg="",this.ended=!1,this.chunks=[],this.strm=new n,this.strm.avail_out=0;var r=d.inflateInit2(this.strm,e.windowBits);if(r!==m.Z_OK)throw new Error(i[r]);this.header=new s,d.inflateGetHeader(this.strm,this.header)}function o(t,e){var r=new a(e);if(r.push(t,!0),r.err)throw r.msg||i[r.err];return r.result}a.prototype.push=function(t,e){var r,i,n,s,a,o,h=this.strm,u=this.options.chunkSize,l=this.options.dictionary,f=!1;if(this.ended)return!1;i=e===~~e?e:!0===e?m.Z_FINISH:m.Z_NO_FLUSH,"string"==typeof t?h.input=p.binstring2buf(t):"[object ArrayBuffer]"===_.call(t)?h.input=new Uint8Array(t):h.input=t,h.next_in=0,h.avail_in=h.input.length;do{if(0===h.avail_out&&(h.output=new c.Buf8(u),h.next_out=0,h.avail_out=u),(r=d.inflate(h,m.Z_NO_FLUSH))===m.Z_NEED_DICT&&l&&(o="string"==typeof l?p.string2buf(l):"[object ArrayBuffer]"===_.call(l)?new Uint8Array(l):l,r=d.inflateSetDictionary(this.strm,o)),r===m.Z_BUF_ERROR&&!0===f&&(r=m.Z_OK,f=!1),r!==m.Z_STREAM_END&&r!==m.Z_OK)return this.onEnd(r),!(this.ended=!0);h.next_out&&(0!==h.avail_out&&r!==m.Z_STREAM_END&&(0!==h.avail_in||i!==m.Z_FINISH&&i!==m.Z_SYNC_FLUSH)||("string"===this.options.to?(n=p.utf8border(h.output,h.next_out),s=h.next_out-n,a=p.buf2string(h.output,n),h.next_out=s,h.avail_out=u-s,s&&c.arraySet(h.output,h.output,n,s,0),this.onData(a)):this.onData(c.shrinkBuf(h.output,h.next_out)))),0===h.avail_in&&0===h.avail_out&&(f=!0)}while((0<h.avail_in||0===h.avail_out)&&r!==m.Z_STREAM_END);return r===m.Z_STREAM_END&&(i=m.Z_FINISH),i===m.Z_FINISH?(r=d.inflateEnd(this.strm),this.onEnd(r),this.ended=!0,r===m.Z_OK):i!==m.Z_SYNC_FLUSH||(this.onEnd(m.Z_OK),!(h.avail_out=0))},a.prototype.onData=function(t){this.chunks.push(t)},a.prototype.onEnd=function(t){t===m.Z_OK&&("string"===this.options.to?this.result=this.chunks.join(""):this.result=c.flattenChunks(this.chunks)),this.chunks=[],this.err=t,this.msg=this.strm.msg},r.Inflate=a,r.inflate=o,r.inflateRaw=function(t,e){return(e=e||{}).raw=!0,o(t,e)},r.ungzip=o},{"./utils/common":41,"./utils/strings":42,"./zlib/constants":44,"./zlib/gzheader":47,"./zlib/inflate":49,"./zlib/messages":51,"./zlib/zstream":53}],41:[function(t,e,r){"use strict";var i="undefined"!=typeof Uint8Array&&"undefined"!=typeof Uint16Array&&"undefined"!=typeof Int32Array;r.assign=function(t){for(var e=Array.prototype.slice.call(arguments,1);e.length;){var r=e.shift();if(r){if("object"!=typeof r)throw new TypeError(r+"must be non-object");for(var i in r)r.hasOwnProperty(i)&&(t[i]=r[i])}}return t},r.shrinkBuf=function(t,e){return t.length===e?t:t.subarray?t.subarray(0,e):(t.length=e,t)};var n={arraySet:function(t,e,r,i,n){if(e.subarray&&t.subarray)t.set(e.subarray(r,r+i),n);else for(var s=0;s<i;s++)t[n+s]=e[r+s]},flattenChunks:function(t){var e,r,i,n,s,a;for(e=i=0,r=t.length;e<r;e++)i+=t[e].length;for(a=new Uint8Array(i),e=n=0,r=t.length;e<r;e++)s=t[e],a.set(s,n),n+=s.length;return a}},s={arraySet:function(t,e,r,i,n){for(var s=0;s<i;s++)t[n+s]=e[r+s]},flattenChunks:function(t){return[].concat.apply([],t)}};r.setTyped=function(t){t?(r.Buf8=Uint8Array,r.Buf16=Uint16Array,r.Buf32=Int32Array,r.assign(r,n)):(r.Buf8=Array,r.Buf16=Array,r.Buf32=Array,r.assign(r,s))},r.setTyped(i)},{}],42:[function(t,e,r){"use strict";var h=t("./common"),n=!0,s=!0;try{String.fromCharCode.apply(null,[0])}catch(t){n=!1}try{String.fromCharCode.apply(null,new Uint8Array(1))}catch(t){s=!1}for(var u=new h.Buf8(256),i=0;i<256;i++)u[i]=252<=i?6:248<=i?5:240<=i?4:224<=i?3:192<=i?2:1;function l(t,e){if(e<65537&&(t.subarray&&s||!t.subarray&&n))return String.fromCharCode.apply(null,h.shrinkBuf(t,e));for(var r="",i=0;i<e;i++)r+=String.fromCharCode(t[i]);return r}u[254]=u[254]=1,r.string2buf=function(t){var e,r,i,n,s,a=t.length,o=0;for(n=0;n<a;n++)55296==(64512&(r=t.charCodeAt(n)))&&n+1<a&&56320==(64512&(i=t.charCodeAt(n+1)))&&(r=65536+(r-55296<<10)+(i-56320),n++),o+=r<128?1:r<2048?2:r<65536?3:4;for(e=new h.Buf8(o),n=s=0;s<o;n++)55296==(64512&(r=t.charCodeAt(n)))&&n+1<a&&56320==(64512&(i=t.charCodeAt(n+1)))&&(r=65536+(r-55296<<10)+(i-56320),n++),r<128?e[s++]=r:(r<2048?e[s++]=192|r>>>6:(r<65536?e[s++]=224|r>>>12:(e[s++]=240|r>>>18,e[s++]=128|r>>>12&63),e[s++]=128|r>>>6&63),e[s++]=128|63&r);return e},r.buf2binstring=function(t){return l(t,t.length)},r.binstring2buf=function(t){for(var e=new h.Buf8(t.length),r=0,i=e.length;r<i;r++)e[r]=t.charCodeAt(r);return e},r.buf2string=function(t,e){var r,i,n,s,a=e||t.length,o=new Array(2*a);for(r=i=0;r<a;)if((n=t[r++])<128)o[i++]=n;else if(4<(s=u[n]))o[i++]=65533,r+=s-1;else{for(n&=2===s?31:3===s?15:7;1<s&&r<a;)n=n<<6|63&t[r++],s--;1<s?o[i++]=65533:n<65536?o[i++]=n:(n-=65536,o[i++]=55296|n>>10&1023,o[i++]=56320|1023&n)}return l(o,i)},r.utf8border=function(t,e){var r;for((e=e||t.length)>t.length&&(e=t.length),r=e-1;0<=r&&128==(192&t[r]);)r--;return r<0?e:0===r?e:r+u[t[r]]>e?r:e}},{"./common":41}],43:[function(t,e,r){"use strict";e.exports=function(t,e,r,i){for(var n=65535&t|0,s=t>>>16&65535|0,a=0;0!==r;){for(r-=a=2e3<r?2e3:r;s=s+(n=n+e[i++]|0)|0,--a;);n%=65521,s%=65521}return n|s<<16|0}},{}],44:[function(t,e,r){"use strict";e.exports={Z_NO_FLUSH:0,Z_PARTIAL_FLUSH:1,Z_SYNC_FLUSH:2,Z_FULL_FLUSH:3,Z_FINISH:4,Z_BLOCK:5,Z_TREES:6,Z_OK:0,Z_STREAM_END:1,Z_NEED_DICT:2,Z_ERRNO:-1,Z_STREAM_ERROR:-2,Z_DATA_ERROR:-3,Z_BUF_ERROR:-5,Z_NO_COMPRESSION:0,Z_BEST_SPEED:1,Z_BEST_COMPRESSION:9,Z_DEFAULT_COMPRESSION:-1,Z_FILTERED:1,Z_HUFFMAN_ONLY:2,Z_RLE:3,Z_FIXED:4,Z_DEFAULT_STRATEGY:0,Z_BINARY:0,Z_TEXT:1,Z_UNKNOWN:2,Z_DEFLATED:8}},{}],45:[function(t,e,r){"use strict";var o=function(){for(var t,e=[],r=0;r<256;r++){t=r;for(var i=0;i<8;i++)t=1&t?3988292384^t>>>1:t>>>1;e[r]=t}return e}();e.exports=function(t,e,r,i){var n=o,s=i+r;t^=-1;for(var a=i;a<s;a++)t=t>>>8^n[255&(t^e[a])];return-1^t}},{}],46:[function(t,e,r){"use strict";var h,d=t("../utils/common"),u=t("./trees"),c=t("./adler32"),p=t("./crc32"),i=t("./messages"),l=0,f=4,m=0,_=-2,g=-1,b=4,n=2,v=8,y=9,s=286,a=30,o=19,w=2*s+1,k=15,x=3,S=258,z=S+x+1,C=42,E=113,A=1,I=2,O=3,B=4;function R(t,e){return t.msg=i[e],e}function T(t){return(t<<1)-(4<t?9:0)}function D(t){for(var e=t.length;0<=--e;)t[e]=0}function F(t){var e=t.state,r=e.pending;r>t.avail_out&&(r=t.avail_out),0!==r&&(d.arraySet(t.output,e.pending_buf,e.pending_out,r,t.next_out),t.next_out+=r,e.pending_out+=r,t.total_out+=r,t.avail_out-=r,e.pending-=r,0===e.pending&&(e.pending_out=0))}function N(t,e){u._tr_flush_block(t,0<=t.block_start?t.block_start:-1,t.strstart-t.block_start,e),t.block_start=t.strstart,F(t.strm)}function U(t,e){t.pending_buf[t.pending++]=e}function P(t,e){t.pending_buf[t.pending++]=e>>>8&255,t.pending_buf[t.pending++]=255&e}function L(t,e){var r,i,n=t.max_chain_length,s=t.strstart,a=t.prev_length,o=t.nice_match,h=t.strstart>t.w_size-z?t.strstart-(t.w_size-z):0,u=t.window,l=t.w_mask,f=t.prev,d=t.strstart+S,c=u[s+a-1],p=u[s+a];t.prev_length>=t.good_match&&(n>>=2),o>t.lookahead&&(o=t.lookahead);do{if(u[(r=e)+a]===p&&u[r+a-1]===c&&u[r]===u[s]&&u[++r]===u[s+1]){s+=2,r++;do{}while(u[++s]===u[++r]&&u[++s]===u[++r]&&u[++s]===u[++r]&&u[++s]===u[++r]&&u[++s]===u[++r]&&u[++s]===u[++r]&&u[++s]===u[++r]&&u[++s]===u[++r]&&s<d);if(i=S-(d-s),s=d-S,a<i){if(t.match_start=e,o<=(a=i))break;c=u[s+a-1],p=u[s+a]}}}while((e=f[e&l])>h&&0!=--n);return a<=t.lookahead?a:t.lookahead}function j(t){var e,r,i,n,s,a,o,h,u,l,f=t.w_size;do{if(n=t.window_size-t.lookahead-t.strstart,t.strstart>=f+(f-z)){for(d.arraySet(t.window,t.window,f,f,0),t.match_start-=f,t.strstart-=f,t.block_start-=f,e=r=t.hash_size;i=t.head[--e],t.head[e]=f<=i?i-f:0,--r;);for(e=r=f;i=t.prev[--e],t.prev[e]=f<=i?i-f:0,--r;);n+=f}if(0===t.strm.avail_in)break;if(a=t.strm,o=t.window,h=t.strstart+t.lookahead,u=n,l=void 0,l=a.avail_in,u<l&&(l=u),r=0===l?0:(a.avail_in-=l,d.arraySet(o,a.input,a.next_in,l,h),1===a.state.wrap?a.adler=c(a.adler,o,l,h):2===a.state.wrap&&(a.adler=p(a.adler,o,l,h)),a.next_in+=l,a.total_in+=l,l),t.lookahead+=r,t.lookahead+t.insert>=x)for(s=t.strstart-t.insert,t.ins_h=t.window[s],t.ins_h=(t.ins_h<<t.hash_shift^t.window[s+1])&t.hash_mask;t.insert&&(t.ins_h=(t.ins_h<<t.hash_shift^t.window[s+x-1])&t.hash_mask,t.prev[s&t.w_mask]=t.head[t.ins_h],t.head[t.ins_h]=s,s++,t.insert--,!(t.lookahead+t.insert<x)););}while(t.lookahead<z&&0!==t.strm.avail_in)}function Z(t,e){for(var r,i;;){if(t.lookahead<z){if(j(t),t.lookahead<z&&e===l)return A;if(0===t.lookahead)break}if(r=0,t.lookahead>=x&&(t.ins_h=(t.ins_h<<t.hash_shift^t.window[t.strstart+x-1])&t.hash_mask,r=t.prev[t.strstart&t.w_mask]=t.head[t.ins_h],t.head[t.ins_h]=t.strstart),0!==r&&t.strstart-r<=t.w_size-z&&(t.match_length=L(t,r)),t.match_length>=x)if(i=u._tr_tally(t,t.strstart-t.match_start,t.match_length-x),t.lookahead-=t.match_length,t.match_length<=t.max_lazy_match&&t.lookahead>=x){for(t.match_length--;t.strstart++,t.ins_h=(t.ins_h<<t.hash_shift^t.window[t.strstart+x-1])&t.hash_mask,r=t.prev[t.strstart&t.w_mask]=t.head[t.ins_h],t.head[t.ins_h]=t.strstart,0!=--t.match_length;);t.strstart++}else t.strstart+=t.match_length,t.match_length=0,t.ins_h=t.window[t.strstart],t.ins_h=(t.ins_h<<t.hash_shift^t.window[t.strstart+1])&t.hash_mask;else i=u._tr_tally(t,0,t.window[t.strstart]),t.lookahead--,t.strstart++;if(i&&(N(t,!1),0===t.strm.avail_out))return A}return t.insert=t.strstart<x-1?t.strstart:x-1,e===f?(N(t,!0),0===t.strm.avail_out?O:B):t.last_lit&&(N(t,!1),0===t.strm.avail_out)?A:I}function W(t,e){for(var r,i,n;;){if(t.lookahead<z){if(j(t),t.lookahead<z&&e===l)return A;if(0===t.lookahead)break}if(r=0,t.lookahead>=x&&(t.ins_h=(t.ins_h<<t.hash_shift^t.window[t.strstart+x-1])&t.hash_mask,r=t.prev[t.strstart&t.w_mask]=t.head[t.ins_h],t.head[t.ins_h]=t.strstart),t.prev_length=t.match_length,t.prev_match=t.match_start,t.match_length=x-1,0!==r&&t.prev_length<t.max_lazy_match&&t.strstart-r<=t.w_size-z&&(t.match_length=L(t,r),t.match_length<=5&&(1===t.strategy||t.match_length===x&&4096<t.strstart-t.match_start)&&(t.match_length=x-1)),t.prev_length>=x&&t.match_length<=t.prev_length){for(n=t.strstart+t.lookahead-x,i=u._tr_tally(t,t.strstart-1-t.prev_match,t.prev_length-x),t.lookahead-=t.prev_length-1,t.prev_length-=2;++t.strstart<=n&&(t.ins_h=(t.ins_h<<t.hash_shift^t.window[t.strstart+x-1])&t.hash_mask,r=t.prev[t.strstart&t.w_mask]=t.head[t.ins_h],t.head[t.ins_h]=t.strstart),0!=--t.prev_length;);if(t.match_available=0,t.match_length=x-1,t.strstart++,i&&(N(t,!1),0===t.strm.avail_out))return A}else if(t.match_available){if((i=u._tr_tally(t,0,t.window[t.strstart-1]))&&N(t,!1),t.strstart++,t.lookahead--,0===t.strm.avail_out)return A}else t.match_available=1,t.strstart++,t.lookahead--}return t.match_available&&(i=u._tr_tally(t,0,t.window[t.strstart-1]),t.match_available=0),t.insert=t.strstart<x-1?t.strstart:x-1,e===f?(N(t,!0),0===t.strm.avail_out?O:B):t.last_lit&&(N(t,!1),0===t.strm.avail_out)?A:I}function M(t,e,r,i,n){this.good_length=t,this.max_lazy=e,this.nice_length=r,this.max_chain=i,this.func=n}function H(){this.strm=null,this.status=0,this.pending_buf=null,this.pending_buf_size=0,this.pending_out=0,this.pending=0,this.wrap=0,this.gzhead=null,this.gzindex=0,this.method=v,this.last_flush=-1,this.w_size=0,this.w_bits=0,this.w_mask=0,this.window=null,this.window_size=0,this.prev=null,this.head=null,this.ins_h=0,this.hash_size=0,this.hash_bits=0,this.hash_mask=0,this.hash_shift=0,this.block_start=0,this.match_length=0,this.prev_match=0,this.match_available=0,this.strstart=0,this.match_start=0,this.lookahead=0,this.prev_length=0,this.max_chain_length=0,this.max_lazy_match=0,this.level=0,this.strategy=0,this.good_match=0,this.nice_match=0,this.dyn_ltree=new d.Buf16(2*w),this.dyn_dtree=new d.Buf16(2*(2*a+1)),this.bl_tree=new d.Buf16(2*(2*o+1)),D(this.dyn_ltree),D(this.dyn_dtree),D(this.bl_tree),this.l_desc=null,this.d_desc=null,this.bl_desc=null,this.bl_count=new d.Buf16(k+1),this.heap=new d.Buf16(2*s+1),D(this.heap),this.heap_len=0,this.heap_max=0,this.depth=new d.Buf16(2*s+1),D(this.depth),this.l_buf=0,this.lit_bufsize=0,this.last_lit=0,this.d_buf=0,this.opt_len=0,this.static_len=0,this.matches=0,this.insert=0,this.bi_buf=0,this.bi_valid=0}function G(t){var e;return t&&t.state?(t.total_in=t.total_out=0,t.data_type=n,(e=t.state).pending=0,e.pending_out=0,e.wrap<0&&(e.wrap=-e.wrap),e.status=e.wrap?C:E,t.adler=2===e.wrap?0:1,e.last_flush=l,u._tr_init(e),m):R(t,_)}function K(t){var e=G(t);return e===m&&function(t){t.window_size=2*t.w_size,D(t.head),t.max_lazy_match=h[t.level].max_lazy,t.good_match=h[t.level].good_length,t.nice_match=h[t.level].nice_length,t.max_chain_length=h[t.level].max_chain,t.strstart=0,t.block_start=0,t.lookahead=0,t.insert=0,t.match_length=t.prev_length=x-1,t.match_available=0,t.ins_h=0}(t.state),e}function Y(t,e,r,i,n,s){if(!t)return _;var a=1;if(e===g&&(e=6),i<0?(a=0,i=-i):15<i&&(a=2,i-=16),n<1||y<n||r!==v||i<8||15<i||e<0||9<e||s<0||b<s)return R(t,_);8===i&&(i=9);var o=new H;return(t.state=o).strm=t,o.wrap=a,o.gzhead=null,o.w_bits=i,o.w_size=1<<o.w_bits,o.w_mask=o.w_size-1,o.hash_bits=n+7,o.hash_size=1<<o.hash_bits,o.hash_mask=o.hash_size-1,o.hash_shift=~~((o.hash_bits+x-1)/x),o.window=new d.Buf8(2*o.w_size),o.head=new d.Buf16(o.hash_size),o.prev=new d.Buf16(o.w_size),o.lit_bufsize=1<<n+6,o.pending_buf_size=4*o.lit_bufsize,o.pending_buf=new d.Buf8(o.pending_buf_size),o.d_buf=1*o.lit_bufsize,o.l_buf=3*o.lit_bufsize,o.level=e,o.strategy=s,o.method=r,K(t)}h=[new M(0,0,0,0,function(t,e){var r=65535;for(r>t.pending_buf_size-5&&(r=t.pending_buf_size-5);;){if(t.lookahead<=1){if(j(t),0===t.lookahead&&e===l)return A;if(0===t.lookahead)break}t.strstart+=t.lookahead,t.lookahead=0;var i=t.block_start+r;if((0===t.strstart||t.strstart>=i)&&(t.lookahead=t.strstart-i,t.strstart=i,N(t,!1),0===t.strm.avail_out))return A;if(t.strstart-t.block_start>=t.w_size-z&&(N(t,!1),0===t.strm.avail_out))return A}return t.insert=0,e===f?(N(t,!0),0===t.strm.avail_out?O:B):(t.strstart>t.block_start&&(N(t,!1),t.strm.avail_out),A)}),new M(4,4,8,4,Z),new M(4,5,16,8,Z),new M(4,6,32,32,Z),new M(4,4,16,16,W),new M(8,16,32,32,W),new M(8,16,128,128,W),new M(8,32,128,256,W),new M(32,128,258,1024,W),new M(32,258,258,4096,W)],r.deflateInit=function(t,e){return Y(t,e,v,15,8,0)},r.deflateInit2=Y,r.deflateReset=K,r.deflateResetKeep=G,r.deflateSetHeader=function(t,e){return t&&t.state?2!==t.state.wrap?_:(t.state.gzhead=e,m):_},r.deflate=function(t,e){var r,i,n,s;if(!t||!t.state||5<e||e<0)return t?R(t,_):_;if(i=t.state,!t.output||!t.input&&0!==t.avail_in||666===i.status&&e!==f)return R(t,0===t.avail_out?-5:_);if(i.strm=t,r=i.last_flush,i.last_flush=e,i.status===C)if(2===i.wrap)t.adler=0,U(i,31),U(i,139),U(i,8),i.gzhead?(U(i,(i.gzhead.text?1:0)+(i.gzhead.hcrc?2:0)+(i.gzhead.extra?4:0)+(i.gzhead.name?8:0)+(i.gzhead.comment?16:0)),U(i,255&i.gzhead.time),U(i,i.gzhead.time>>8&255),U(i,i.gzhead.time>>16&255),U(i,i.gzhead.time>>24&255),U(i,9===i.level?2:2<=i.strategy||i.level<2?4:0),U(i,255&i.gzhead.os),i.gzhead.extra&&i.gzhead.extra.length&&(U(i,255&i.gzhead.extra.length),U(i,i.gzhead.extra.length>>8&255)),i.gzhead.hcrc&&(t.adler=p(t.adler,i.pending_buf,i.pending,0)),i.gzindex=0,i.status=69):(U(i,0),U(i,0),U(i,0),U(i,0),U(i,0),U(i,9===i.level?2:2<=i.strategy||i.level<2?4:0),U(i,3),i.status=E);else{var a=v+(i.w_bits-8<<4)<<8;a|=(2<=i.strategy||i.level<2?0:i.level<6?1:6===i.level?2:3)<<6,0!==i.strstart&&(a|=32),a+=31-a%31,i.status=E,P(i,a),0!==i.strstart&&(P(i,t.adler>>>16),P(i,65535&t.adler)),t.adler=1}if(69===i.status)if(i.gzhead.extra){for(n=i.pending;i.gzindex<(65535&i.gzhead.extra.length)&&(i.pending!==i.pending_buf_size||(i.gzhead.hcrc&&i.pending>n&&(t.adler=p(t.adler,i.pending_buf,i.pending-n,n)),F(t),n=i.pending,i.pending!==i.pending_buf_size));)U(i,255&i.gzhead.extra[i.gzindex]),i.gzindex++;i.gzhead.hcrc&&i.pending>n&&(t.adler=p(t.adler,i.pending_buf,i.pending-n,n)),i.gzindex===i.gzhead.extra.length&&(i.gzindex=0,i.status=73)}else i.status=73;if(73===i.status)if(i.gzhead.name){n=i.pending;do{if(i.pending===i.pending_buf_size&&(i.gzhead.hcrc&&i.pending>n&&(t.adler=p(t.adler,i.pending_buf,i.pending-n,n)),F(t),n=i.pending,i.pending===i.pending_buf_size)){s=1;break}s=i.gzindex<i.gzhead.name.length?255&i.gzhead.name.charCodeAt(i.gzindex++):0,U(i,s)}while(0!==s);i.gzhead.hcrc&&i.pending>n&&(t.adler=p(t.adler,i.pending_buf,i.pending-n,n)),0===s&&(i.gzindex=0,i.status=91)}else i.status=91;if(91===i.status)if(i.gzhead.comment){n=i.pending;do{if(i.pending===i.pending_buf_size&&(i.gzhead.hcrc&&i.pending>n&&(t.adler=p(t.adler,i.pending_buf,i.pending-n,n)),F(t),n=i.pending,i.pending===i.pending_buf_size)){s=1;break}s=i.gzindex<i.gzhead.comment.length?255&i.gzhead.comment.charCodeAt(i.gzindex++):0,U(i,s)}while(0!==s);i.gzhead.hcrc&&i.pending>n&&(t.adler=p(t.adler,i.pending_buf,i.pending-n,n)),0===s&&(i.status=103)}else i.status=103;if(103===i.status&&(i.gzhead.hcrc?(i.pending+2>i.pending_buf_size&&F(t),i.pending+2<=i.pending_buf_size&&(U(i,255&t.adler),U(i,t.adler>>8&255),t.adler=0,i.status=E)):i.status=E),0!==i.pending){if(F(t),0===t.avail_out)return i.last_flush=-1,m}else if(0===t.avail_in&&T(e)<=T(r)&&e!==f)return R(t,-5);if(666===i.status&&0!==t.avail_in)return R(t,-5);if(0!==t.avail_in||0!==i.lookahead||e!==l&&666!==i.status){var o=2===i.strategy?function(t,e){for(var r;;){if(0===t.lookahead&&(j(t),0===t.lookahead)){if(e===l)return A;break}if(t.match_length=0,r=u._tr_tally(t,0,t.window[t.strstart]),t.lookahead--,t.strstart++,r&&(N(t,!1),0===t.strm.avail_out))return A}return t.insert=0,e===f?(N(t,!0),0===t.strm.avail_out?O:B):t.last_lit&&(N(t,!1),0===t.strm.avail_out)?A:I}(i,e):3===i.strategy?function(t,e){for(var r,i,n,s,a=t.window;;){if(t.lookahead<=S){if(j(t),t.lookahead<=S&&e===l)return A;if(0===t.lookahead)break}if(t.match_length=0,t.lookahead>=x&&0<t.strstart&&(i=a[n=t.strstart-1])===a[++n]&&i===a[++n]&&i===a[++n]){s=t.strstart+S;do{}while(i===a[++n]&&i===a[++n]&&i===a[++n]&&i===a[++n]&&i===a[++n]&&i===a[++n]&&i===a[++n]&&i===a[++n]&&n<s);t.match_length=S-(s-n),t.match_length>t.lookahead&&(t.match_length=t.lookahead)}if(t.match_length>=x?(r=u._tr_tally(t,1,t.match_length-x),t.lookahead-=t.match_length,t.strstart+=t.match_length,t.match_length=0):(r=u._tr_tally(t,0,t.window[t.strstart]),t.lookahead--,t.strstart++),r&&(N(t,!1),0===t.strm.avail_out))return A}return t.insert=0,e===f?(N(t,!0),0===t.strm.avail_out?O:B):t.last_lit&&(N(t,!1),0===t.strm.avail_out)?A:I}(i,e):h[i.level].func(i,e);if(o!==O&&o!==B||(i.status=666),o===A||o===O)return 0===t.avail_out&&(i.last_flush=-1),m;if(o===I&&(1===e?u._tr_align(i):5!==e&&(u._tr_stored_block(i,0,0,!1),3===e&&(D(i.head),0===i.lookahead&&(i.strstart=0,i.block_start=0,i.insert=0))),F(t),0===t.avail_out))return i.last_flush=-1,m}return e!==f?m:i.wrap<=0?1:(2===i.wrap?(U(i,255&t.adler),U(i,t.adler>>8&255),U(i,t.adler>>16&255),U(i,t.adler>>24&255),U(i,255&t.total_in),U(i,t.total_in>>8&255),U(i,t.total_in>>16&255),U(i,t.total_in>>24&255)):(P(i,t.adler>>>16),P(i,65535&t.adler)),F(t),0<i.wrap&&(i.wrap=-i.wrap),0!==i.pending?m:1)},r.deflateEnd=function(t){var e;return t&&t.state?(e=t.state.status)!==C&&69!==e&&73!==e&&91!==e&&103!==e&&e!==E&&666!==e?R(t,_):(t.state=null,e===E?R(t,-3):m):_},r.deflateSetDictionary=function(t,e){var r,i,n,s,a,o,h,u,l=e.length;if(!t||!t.state)return _;if(2===(s=(r=t.state).wrap)||1===s&&r.status!==C||r.lookahead)return _;for(1===s&&(t.adler=c(t.adler,e,l,0)),r.wrap=0,l>=r.w_size&&(0===s&&(D(r.head),r.strstart=0,r.block_start=0,r.insert=0),u=new d.Buf8(r.w_size),d.arraySet(u,e,l-r.w_size,r.w_size,0),e=u,l=r.w_size),a=t.avail_in,o=t.next_in,h=t.input,t.avail_in=l,t.next_in=0,t.input=e,j(r);r.lookahead>=x;){for(i=r.strstart,n=r.lookahead-(x-1);r.ins_h=(r.ins_h<<r.hash_shift^r.window[i+x-1])&r.hash_mask,r.prev[i&r.w_mask]=r.head[r.ins_h],r.head[r.ins_h]=i,i++,--n;);r.strstart=i,r.lookahead=x-1,j(r)}return r.strstart+=r.lookahead,r.block_start=r.strstart,r.insert=r.lookahead,r.lookahead=0,r.match_length=r.prev_length=x-1,r.match_available=0,t.next_in=o,t.input=h,t.avail_in=a,r.wrap=s,m},r.deflateInfo="pako deflate (from Nodeca project)"},{"../utils/common":41,"./adler32":43,"./crc32":45,"./messages":51,"./trees":52}],47:[function(t,e,r){"use strict";e.exports=function(){this.text=0,this.time=0,this.xflags=0,this.os=0,this.extra=null,this.extra_len=0,this.name="",this.comment="",this.hcrc=0,this.done=!1}},{}],48:[function(t,e,r){"use strict";e.exports=function(t,e){var r,i,n,s,a,o,h,u,l,f,d,c,p,m,_,g,b,v,y,w,k,x,S,z,C;r=t.state,i=t.next_in,z=t.input,n=i+(t.avail_in-5),s=t.next_out,C=t.output,a=s-(e-t.avail_out),o=s+(t.avail_out-257),h=r.dmax,u=r.wsize,l=r.whave,f=r.wnext,d=r.window,c=r.hold,p=r.bits,m=r.lencode,_=r.distcode,g=(1<<r.lenbits)-1,b=(1<<r.distbits)-1;t:do{p<15&&(c+=z[i++]<<p,p+=8,c+=z[i++]<<p,p+=8),v=m[c&g];e:for(;;){if(c>>>=y=v>>>24,p-=y,0===(y=v>>>16&255))C[s++]=65535&v;else{if(!(16&y)){if(0==(64&y)){v=m[(65535&v)+(c&(1<<y)-1)];continue e}if(32&y){r.mode=12;break t}t.msg="invalid literal/length code",r.mode=30;break t}w=65535&v,(y&=15)&&(p<y&&(c+=z[i++]<<p,p+=8),w+=c&(1<<y)-1,c>>>=y,p-=y),p<15&&(c+=z[i++]<<p,p+=8,c+=z[i++]<<p,p+=8),v=_[c&b];r:for(;;){if(c>>>=y=v>>>24,p-=y,!(16&(y=v>>>16&255))){if(0==(64&y)){v=_[(65535&v)+(c&(1<<y)-1)];continue r}t.msg="invalid distance code",r.mode=30;break t}if(k=65535&v,p<(y&=15)&&(c+=z[i++]<<p,(p+=8)<y&&(c+=z[i++]<<p,p+=8)),h<(k+=c&(1<<y)-1)){t.msg="invalid distance too far back",r.mode=30;break t}if(c>>>=y,p-=y,(y=s-a)<k){if(l<(y=k-y)&&r.sane){t.msg="invalid distance too far back",r.mode=30;break t}if(S=d,(x=0)===f){if(x+=u-y,y<w){for(w-=y;C[s++]=d[x++],--y;);x=s-k,S=C}}else if(f<y){if(x+=u+f-y,(y-=f)<w){for(w-=y;C[s++]=d[x++],--y;);if(x=0,f<w){for(w-=y=f;C[s++]=d[x++],--y;);x=s-k,S=C}}}else if(x+=f-y,y<w){for(w-=y;C[s++]=d[x++],--y;);x=s-k,S=C}for(;2<w;)C[s++]=S[x++],C[s++]=S[x++],C[s++]=S[x++],w-=3;w&&(C[s++]=S[x++],1<w&&(C[s++]=S[x++]))}else{for(x=s-k;C[s++]=C[x++],C[s++]=C[x++],C[s++]=C[x++],2<(w-=3););w&&(C[s++]=C[x++],1<w&&(C[s++]=C[x++]))}break}}break}}while(i<n&&s<o);i-=w=p>>3,c&=(1<<(p-=w<<3))-1,t.next_in=i,t.next_out=s,t.avail_in=i<n?n-i+5:5-(i-n),t.avail_out=s<o?o-s+257:257-(s-o),r.hold=c,r.bits=p}},{}],49:[function(t,e,r){"use strict";var I=t("../utils/common"),O=t("./adler32"),B=t("./crc32"),R=t("./inffast"),T=t("./inftrees"),D=1,F=2,N=0,U=-2,P=1,i=852,n=592;function L(t){return(t>>>24&255)+(t>>>8&65280)+((65280&t)<<8)+((255&t)<<24)}function s(){this.mode=0,this.last=!1,this.wrap=0,this.havedict=!1,this.flags=0,this.dmax=0,this.check=0,this.total=0,this.head=null,this.wbits=0,this.wsize=0,this.whave=0,this.wnext=0,this.window=null,this.hold=0,this.bits=0,this.length=0,this.offset=0,this.extra=0,this.lencode=null,this.distcode=null,this.lenbits=0,this.distbits=0,this.ncode=0,this.nlen=0,this.ndist=0,this.have=0,this.next=null,this.lens=new I.Buf16(320),this.work=new I.Buf16(288),this.lendyn=null,this.distdyn=null,this.sane=0,this.back=0,this.was=0}function a(t){var e;return t&&t.state?(e=t.state,t.total_in=t.total_out=e.total=0,t.msg="",e.wrap&&(t.adler=1&e.wrap),e.mode=P,e.last=0,e.havedict=0,e.dmax=32768,e.head=null,e.hold=0,e.bits=0,e.lencode=e.lendyn=new I.Buf32(i),e.distcode=e.distdyn=new I.Buf32(n),e.sane=1,e.back=-1,N):U}function o(t){var e;return t&&t.state?((e=t.state).wsize=0,e.whave=0,e.wnext=0,a(t)):U}function h(t,e){var r,i;return t&&t.state?(i=t.state,e<0?(r=0,e=-e):(r=1+(e>>4),e<48&&(e&=15)),e&&(e<8||15<e)?U:(null!==i.window&&i.wbits!==e&&(i.window=null),i.wrap=r,i.wbits=e,o(t))):U}function u(t,e){var r,i;return t?(i=new s,(t.state=i).window=null,(r=h(t,e))!==N&&(t.state=null),r):U}var l,f,d=!0;function j(t){if(d){var e;for(l=new I.Buf32(512),f=new I.Buf32(32),e=0;e<144;)t.lens[e++]=8;for(;e<256;)t.lens[e++]=9;for(;e<280;)t.lens[e++]=7;for(;e<288;)t.lens[e++]=8;for(T(D,t.lens,0,288,l,0,t.work,{bits:9}),e=0;e<32;)t.lens[e++]=5;T(F,t.lens,0,32,f,0,t.work,{bits:5}),d=!1}t.lencode=l,t.lenbits=9,t.distcode=f,t.distbits=5}function Z(t,e,r,i){var n,s=t.state;return null===s.window&&(s.wsize=1<<s.wbits,s.wnext=0,s.whave=0,s.window=new I.Buf8(s.wsize)),i>=s.wsize?(I.arraySet(s.window,e,r-s.wsize,s.wsize,0),s.wnext=0,s.whave=s.wsize):(i<(n=s.wsize-s.wnext)&&(n=i),I.arraySet(s.window,e,r-i,n,s.wnext),(i-=n)?(I.arraySet(s.window,e,r-i,i,0),s.wnext=i,s.whave=s.wsize):(s.wnext+=n,s.wnext===s.wsize&&(s.wnext=0),s.whave<s.wsize&&(s.whave+=n))),0}r.inflateReset=o,r.inflateReset2=h,r.inflateResetKeep=a,r.inflateInit=function(t){return u(t,15)},r.inflateInit2=u,r.inflate=function(t,e){var r,i,n,s,a,o,h,u,l,f,d,c,p,m,_,g,b,v,y,w,k,x,S,z,C=0,E=new I.Buf8(4),A=[16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15];if(!t||!t.state||!t.output||!t.input&&0!==t.avail_in)return U;12===(r=t.state).mode&&(r.mode=13),a=t.next_out,n=t.output,h=t.avail_out,s=t.next_in,i=t.input,o=t.avail_in,u=r.hold,l=r.bits,f=o,d=h,x=N;t:for(;;)switch(r.mode){case P:if(0===r.wrap){r.mode=13;break}for(;l<16;){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}if(2&r.wrap&&35615===u){E[r.check=0]=255&u,E[1]=u>>>8&255,r.check=B(r.check,E,2,0),l=u=0,r.mode=2;break}if(r.flags=0,r.head&&(r.head.done=!1),!(1&r.wrap)||(((255&u)<<8)+(u>>8))%31){t.msg="incorrect header check",r.mode=30;break}if(8!=(15&u)){t.msg="unknown compression method",r.mode=30;break}if(l-=4,k=8+(15&(u>>>=4)),0===r.wbits)r.wbits=k;else if(k>r.wbits){t.msg="invalid window size",r.mode=30;break}r.dmax=1<<k,t.adler=r.check=1,r.mode=512&u?10:12,l=u=0;break;case 2:for(;l<16;){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}if(r.flags=u,8!=(255&r.flags)){t.msg="unknown compression method",r.mode=30;break}if(57344&r.flags){t.msg="unknown header flags set",r.mode=30;break}r.head&&(r.head.text=u>>8&1),512&r.flags&&(E[0]=255&u,E[1]=u>>>8&255,r.check=B(r.check,E,2,0)),l=u=0,r.mode=3;case 3:for(;l<32;){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}r.head&&(r.head.time=u),512&r.flags&&(E[0]=255&u,E[1]=u>>>8&255,E[2]=u>>>16&255,E[3]=u>>>24&255,r.check=B(r.check,E,4,0)),l=u=0,r.mode=4;case 4:for(;l<16;){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}r.head&&(r.head.xflags=255&u,r.head.os=u>>8),512&r.flags&&(E[0]=255&u,E[1]=u>>>8&255,r.check=B(r.check,E,2,0)),l=u=0,r.mode=5;case 5:if(1024&r.flags){for(;l<16;){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}r.length=u,r.head&&(r.head.extra_len=u),512&r.flags&&(E[0]=255&u,E[1]=u>>>8&255,r.check=B(r.check,E,2,0)),l=u=0}else r.head&&(r.head.extra=null);r.mode=6;case 6:if(1024&r.flags&&(o<(c=r.length)&&(c=o),c&&(r.head&&(k=r.head.extra_len-r.length,r.head.extra||(r.head.extra=new Array(r.head.extra_len)),I.arraySet(r.head.extra,i,s,c,k)),512&r.flags&&(r.check=B(r.check,i,c,s)),o-=c,s+=c,r.length-=c),r.length))break t;r.length=0,r.mode=7;case 7:if(2048&r.flags){if(0===o)break t;for(c=0;k=i[s+c++],r.head&&k&&r.length<65536&&(r.head.name+=String.fromCharCode(k)),k&&c<o;);if(512&r.flags&&(r.check=B(r.check,i,c,s)),o-=c,s+=c,k)break t}else r.head&&(r.head.name=null);r.length=0,r.mode=8;case 8:if(4096&r.flags){if(0===o)break t;for(c=0;k=i[s+c++],r.head&&k&&r.length<65536&&(r.head.comment+=String.fromCharCode(k)),k&&c<o;);if(512&r.flags&&(r.check=B(r.check,i,c,s)),o-=c,s+=c,k)break t}else r.head&&(r.head.comment=null);r.mode=9;case 9:if(512&r.flags){for(;l<16;){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}if(u!==(65535&r.check)){t.msg="header crc mismatch",r.mode=30;break}l=u=0}r.head&&(r.head.hcrc=r.flags>>9&1,r.head.done=!0),t.adler=r.check=0,r.mode=12;break;case 10:for(;l<32;){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}t.adler=r.check=L(u),l=u=0,r.mode=11;case 11:if(0===r.havedict)return t.next_out=a,t.avail_out=h,t.next_in=s,t.avail_in=o,r.hold=u,r.bits=l,2;t.adler=r.check=1,r.mode=12;case 12:if(5===e||6===e)break t;case 13:if(r.last){u>>>=7&l,l-=7&l,r.mode=27;break}for(;l<3;){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}switch(r.last=1&u,l-=1,3&(u>>>=1)){case 0:r.mode=14;break;case 1:if(j(r),r.mode=20,6!==e)break;u>>>=2,l-=2;break t;case 2:r.mode=17;break;case 3:t.msg="invalid block type",r.mode=30}u>>>=2,l-=2;break;case 14:for(u>>>=7&l,l-=7&l;l<32;){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}if((65535&u)!=(u>>>16^65535)){t.msg="invalid stored block lengths",r.mode=30;break}if(r.length=65535&u,l=u=0,r.mode=15,6===e)break t;case 15:r.mode=16;case 16:if(c=r.length){if(o<c&&(c=o),h<c&&(c=h),0===c)break t;I.arraySet(n,i,s,c,a),o-=c,s+=c,h-=c,a+=c,r.length-=c;break}r.mode=12;break;case 17:for(;l<14;){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}if(r.nlen=257+(31&u),u>>>=5,l-=5,r.ndist=1+(31&u),u>>>=5,l-=5,r.ncode=4+(15&u),u>>>=4,l-=4,286<r.nlen||30<r.ndist){t.msg="too many length or distance symbols",r.mode=30;break}r.have=0,r.mode=18;case 18:for(;r.have<r.ncode;){for(;l<3;){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}r.lens[A[r.have++]]=7&u,u>>>=3,l-=3}for(;r.have<19;)r.lens[A[r.have++]]=0;if(r.lencode=r.lendyn,r.lenbits=7,S={bits:r.lenbits},x=T(0,r.lens,0,19,r.lencode,0,r.work,S),r.lenbits=S.bits,x){t.msg="invalid code lengths set",r.mode=30;break}r.have=0,r.mode=19;case 19:for(;r.have<r.nlen+r.ndist;){for(;g=(C=r.lencode[u&(1<<r.lenbits)-1])>>>16&255,b=65535&C,!((_=C>>>24)<=l);){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}if(b<16)u>>>=_,l-=_,r.lens[r.have++]=b;else{if(16===b){for(z=_+2;l<z;){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}if(u>>>=_,l-=_,0===r.have){t.msg="invalid bit length repeat",r.mode=30;break}k=r.lens[r.have-1],c=3+(3&u),u>>>=2,l-=2}else if(17===b){for(z=_+3;l<z;){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}l-=_,k=0,c=3+(7&(u>>>=_)),u>>>=3,l-=3}else{for(z=_+7;l<z;){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}l-=_,k=0,c=11+(127&(u>>>=_)),u>>>=7,l-=7}if(r.have+c>r.nlen+r.ndist){t.msg="invalid bit length repeat",r.mode=30;break}for(;c--;)r.lens[r.have++]=k}}if(30===r.mode)break;if(0===r.lens[256]){t.msg="invalid code -- missing end-of-block",r.mode=30;break}if(r.lenbits=9,S={bits:r.lenbits},x=T(D,r.lens,0,r.nlen,r.lencode,0,r.work,S),r.lenbits=S.bits,x){t.msg="invalid literal/lengths set",r.mode=30;break}if(r.distbits=6,r.distcode=r.distdyn,S={bits:r.distbits},x=T(F,r.lens,r.nlen,r.ndist,r.distcode,0,r.work,S),r.distbits=S.bits,x){t.msg="invalid distances set",r.mode=30;break}if(r.mode=20,6===e)break t;case 20:r.mode=21;case 21:if(6<=o&&258<=h){t.next_out=a,t.avail_out=h,t.next_in=s,t.avail_in=o,r.hold=u,r.bits=l,R(t,d),a=t.next_out,n=t.output,h=t.avail_out,s=t.next_in,i=t.input,o=t.avail_in,u=r.hold,l=r.bits,12===r.mode&&(r.back=-1);break}for(r.back=0;g=(C=r.lencode[u&(1<<r.lenbits)-1])>>>16&255,b=65535&C,!((_=C>>>24)<=l);){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}if(g&&0==(240&g)){for(v=_,y=g,w=b;g=(C=r.lencode[w+((u&(1<<v+y)-1)>>v)])>>>16&255,b=65535&C,!(v+(_=C>>>24)<=l);){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}u>>>=v,l-=v,r.back+=v}if(u>>>=_,l-=_,r.back+=_,r.length=b,0===g){r.mode=26;break}if(32&g){r.back=-1,r.mode=12;break}if(64&g){t.msg="invalid literal/length code",r.mode=30;break}r.extra=15&g,r.mode=22;case 22:if(r.extra){for(z=r.extra;l<z;){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}r.length+=u&(1<<r.extra)-1,u>>>=r.extra,l-=r.extra,r.back+=r.extra}r.was=r.length,r.mode=23;case 23:for(;g=(C=r.distcode[u&(1<<r.distbits)-1])>>>16&255,b=65535&C,!((_=C>>>24)<=l);){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}if(0==(240&g)){for(v=_,y=g,w=b;g=(C=r.distcode[w+((u&(1<<v+y)-1)>>v)])>>>16&255,b=65535&C,!(v+(_=C>>>24)<=l);){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}u>>>=v,l-=v,r.back+=v}if(u>>>=_,l-=_,r.back+=_,64&g){t.msg="invalid distance code",r.mode=30;break}r.offset=b,r.extra=15&g,r.mode=24;case 24:if(r.extra){for(z=r.extra;l<z;){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}r.offset+=u&(1<<r.extra)-1,u>>>=r.extra,l-=r.extra,r.back+=r.extra}if(r.offset>r.dmax){t.msg="invalid distance too far back",r.mode=30;break}r.mode=25;case 25:if(0===h)break t;if(c=d-h,r.offset>c){if((c=r.offset-c)>r.whave&&r.sane){t.msg="invalid distance too far back",r.mode=30;break}p=c>r.wnext?(c-=r.wnext,r.wsize-c):r.wnext-c,c>r.length&&(c=r.length),m=r.window}else m=n,p=a-r.offset,c=r.length;for(h<c&&(c=h),h-=c,r.length-=c;n[a++]=m[p++],--c;);0===r.length&&(r.mode=21);break;case 26:if(0===h)break t;n[a++]=r.length,h--,r.mode=21;break;case 27:if(r.wrap){for(;l<32;){if(0===o)break t;o--,u|=i[s++]<<l,l+=8}if(d-=h,t.total_out+=d,r.total+=d,d&&(t.adler=r.check=r.flags?B(r.check,n,d,a-d):O(r.check,n,d,a-d)),d=h,(r.flags?u:L(u))!==r.check){t.msg="incorrect data check",r.mode=30;break}l=u=0}r.mode=28;case 28:if(r.wrap&&r.flags){for(;l<32;){if(0===o)break t;o--,u+=i[s++]<<l,l+=8}if(u!==(4294967295&r.total)){t.msg="incorrect length check",r.mode=30;break}l=u=0}r.mode=29;case 29:x=1;break t;case 30:x=-3;break t;case 31:return-4;case 32:default:return U}return t.next_out=a,t.avail_out=h,t.next_in=s,t.avail_in=o,r.hold=u,r.bits=l,(r.wsize||d!==t.avail_out&&r.mode<30&&(r.mode<27||4!==e))&&Z(t,t.output,t.next_out,d-t.avail_out)?(r.mode=31,-4):(f-=t.avail_in,d-=t.avail_out,t.total_in+=f,t.total_out+=d,r.total+=d,r.wrap&&d&&(t.adler=r.check=r.flags?B(r.check,n,d,t.next_out-d):O(r.check,n,d,t.next_out-d)),t.data_type=r.bits+(r.last?64:0)+(12===r.mode?128:0)+(20===r.mode||15===r.mode?256:0),(0==f&&0===d||4===e)&&x===N&&(x=-5),x)},r.inflateEnd=function(t){if(!t||!t.state)return U;var e=t.state;return e.window&&(e.window=null),t.state=null,N},r.inflateGetHeader=function(t,e){var r;return t&&t.state?0==(2&(r=t.state).wrap)?U:((r.head=e).done=!1,N):U},r.inflateSetDictionary=function(t,e){var r,i=e.length;return t&&t.state?0!==(r=t.state).wrap&&11!==r.mode?U:11===r.mode&&O(1,e,i,0)!==r.check?-3:Z(t,e,i,i)?(r.mode=31,-4):(r.havedict=1,N):U},r.inflateInfo="pako inflate (from Nodeca project)"},{"../utils/common":41,"./adler32":43,"./crc32":45,"./inffast":48,"./inftrees":50}],50:[function(t,e,r){"use strict";var D=t("../utils/common"),F=[3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258,0,0],N=[16,16,16,16,16,16,16,16,17,17,17,17,18,18,18,18,19,19,19,19,20,20,20,20,21,21,21,21,16,72,78],U=[1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577,0,0],P=[16,16,16,16,17,17,18,18,19,19,20,20,21,21,22,22,23,23,24,24,25,25,26,26,27,27,28,28,29,29,64,64];e.exports=function(t,e,r,i,n,s,a,o){var h,u,l,f,d,c,p,m,_,g=o.bits,b=0,v=0,y=0,w=0,k=0,x=0,S=0,z=0,C=0,E=0,A=null,I=0,O=new D.Buf16(16),B=new D.Buf16(16),R=null,T=0;for(b=0;b<=15;b++)O[b]=0;for(v=0;v<i;v++)O[e[r+v]]++;for(k=g,w=15;1<=w&&0===O[w];w--);if(w<k&&(k=w),0===w)return n[s++]=20971520,n[s++]=20971520,o.bits=1,0;for(y=1;y<w&&0===O[y];y++);for(k<y&&(k=y),b=z=1;b<=15;b++)if(z<<=1,(z-=O[b])<0)return-1;if(0<z&&(0===t||1!==w))return-1;for(B[1]=0,b=1;b<15;b++)B[b+1]=B[b]+O[b];for(v=0;v<i;v++)0!==e[r+v]&&(a[B[e[r+v]]++]=v);if(c=0===t?(A=R=a,19):1===t?(A=F,I-=257,R=N,T-=257,256):(A=U,R=P,-1),b=y,d=s,S=v=E=0,l=-1,f=(C=1<<(x=k))-1,1===t&&852<C||2===t&&592<C)return 1;for(;;){for(p=b-S,_=a[v]<c?(m=0,a[v]):a[v]>c?(m=R[T+a[v]],A[I+a[v]]):(m=96,0),h=1<<b-S,y=u=1<<x;n[d+(E>>S)+(u-=h)]=p<<24|m<<16|_|0,0!==u;);for(h=1<<b-1;E&h;)h>>=1;if(0!==h?(E&=h-1,E+=h):E=0,v++,0==--O[b]){if(b===w)break;b=e[r+a[v]]}if(k<b&&(E&f)!==l){for(0===S&&(S=k),d+=y,z=1<<(x=b-S);x+S<w&&!((z-=O[x+S])<=0);)x++,z<<=1;if(C+=1<<x,1===t&&852<C||2===t&&592<C)return 1;n[l=E&f]=k<<24|x<<16|d-s|0}}return 0!==E&&(n[d+E]=b-S<<24|64<<16|0),o.bits=k,0}},{"../utils/common":41}],51:[function(t,e,r){"use strict";e.exports={2:"need dictionary",1:"stream end",0:"","-1":"file error","-2":"stream error","-3":"data error","-4":"insufficient memory","-5":"buffer error","-6":"incompatible version"}},{}],52:[function(t,e,r){"use strict";var n=t("../utils/common"),o=0,h=1;function i(t){for(var e=t.length;0<=--e;)t[e]=0}var s=0,a=29,u=256,l=u+1+a,f=30,d=19,_=2*l+1,g=15,c=16,p=7,m=256,b=16,v=17,y=18,w=[0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0],k=[0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13],x=[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,2,3,7],S=[16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15],z=new Array(2*(l+2));i(z);var C=new Array(2*f);i(C);var E=new Array(512);i(E);var A=new Array(256);i(A);var I=new Array(a);i(I);var O,B,R,T=new Array(f);function D(t,e,r,i,n){this.static_tree=t,this.extra_bits=e,this.extra_base=r,this.elems=i,this.max_length=n,this.has_stree=t&&t.length}function F(t,e){this.dyn_tree=t,this.max_code=0,this.stat_desc=e}function N(t){return t<256?E[t]:E[256+(t>>>7)]}function U(t,e){t.pending_buf[t.pending++]=255&e,t.pending_buf[t.pending++]=e>>>8&255}function P(t,e,r){t.bi_valid>c-r?(t.bi_buf|=e<<t.bi_valid&65535,U(t,t.bi_buf),t.bi_buf=e>>c-t.bi_valid,t.bi_valid+=r-c):(t.bi_buf|=e<<t.bi_valid&65535,t.bi_valid+=r)}function L(t,e,r){P(t,r[2*e],r[2*e+1])}function j(t,e){for(var r=0;r|=1&t,t>>>=1,r<<=1,0<--e;);return r>>>1}function Z(t,e,r){var i,n,s=new Array(g+1),a=0;for(i=1;i<=g;i++)s[i]=a=a+r[i-1]<<1;for(n=0;n<=e;n++){var o=t[2*n+1];0!==o&&(t[2*n]=j(s[o]++,o))}}function W(t){var e;for(e=0;e<l;e++)t.dyn_ltree[2*e]=0;for(e=0;e<f;e++)t.dyn_dtree[2*e]=0;for(e=0;e<d;e++)t.bl_tree[2*e]=0;t.dyn_ltree[2*m]=1,t.opt_len=t.static_len=0,t.last_lit=t.matches=0}function M(t){8<t.bi_valid?U(t,t.bi_buf):0<t.bi_valid&&(t.pending_buf[t.pending++]=t.bi_buf),t.bi_buf=0,t.bi_valid=0}function H(t,e,r,i){var n=2*e,s=2*r;return t[n]<t[s]||t[n]===t[s]&&i[e]<=i[r]}function G(t,e,r){for(var i=t.heap[r],n=r<<1;n<=t.heap_len&&(n<t.heap_len&&H(e,t.heap[n+1],t.heap[n],t.depth)&&n++,!H(e,i,t.heap[n],t.depth));)t.heap[r]=t.heap[n],r=n,n<<=1;t.heap[r]=i}function K(t,e,r){var i,n,s,a,o=0;if(0!==t.last_lit)for(;i=t.pending_buf[t.d_buf+2*o]<<8|t.pending_buf[t.d_buf+2*o+1],n=t.pending_buf[t.l_buf+o],o++,0===i?L(t,n,e):(L(t,(s=A[n])+u+1,e),0!==(a=w[s])&&P(t,n-=I[s],a),L(t,s=N(--i),r),0!==(a=k[s])&&P(t,i-=T[s],a)),o<t.last_lit;);L(t,m,e)}function Y(t,e){var r,i,n,s=e.dyn_tree,a=e.stat_desc.static_tree,o=e.stat_desc.has_stree,h=e.stat_desc.elems,u=-1;for(t.heap_len=0,t.heap_max=_,r=0;r<h;r++)0!==s[2*r]?(t.heap[++t.heap_len]=u=r,t.depth[r]=0):s[2*r+1]=0;for(;t.heap_len<2;)s[2*(n=t.heap[++t.heap_len]=u<2?++u:0)]=1,t.depth[n]=0,t.opt_len--,o&&(t.static_len-=a[2*n+1]);for(e.max_code=u,r=t.heap_len>>1;1<=r;r--)G(t,s,r);for(n=h;r=t.heap[1],t.heap[1]=t.heap[t.heap_len--],G(t,s,1),i=t.heap[1],t.heap[--t.heap_max]=r,t.heap[--t.heap_max]=i,s[2*n]=s[2*r]+s[2*i],t.depth[n]=(t.depth[r]>=t.depth[i]?t.depth[r]:t.depth[i])+1,s[2*r+1]=s[2*i+1]=n,t.heap[1]=n++,G(t,s,1),2<=t.heap_len;);t.heap[--t.heap_max]=t.heap[1],function(t,e){var r,i,n,s,a,o,h=e.dyn_tree,u=e.max_code,l=e.stat_desc.static_tree,f=e.stat_desc.has_stree,d=e.stat_desc.extra_bits,c=e.stat_desc.extra_base,p=e.stat_desc.max_length,m=0;for(s=0;s<=g;s++)t.bl_count[s]=0;for(h[2*t.heap[t.heap_max]+1]=0,r=t.heap_max+1;r<_;r++)p<(s=h[2*h[2*(i=t.heap[r])+1]+1]+1)&&(s=p,m++),h[2*i+1]=s,u<i||(t.bl_count[s]++,a=0,c<=i&&(a=d[i-c]),o=h[2*i],t.opt_len+=o*(s+a),f&&(t.static_len+=o*(l[2*i+1]+a)));if(0!==m){do{for(s=p-1;0===t.bl_count[s];)s--;t.bl_count[s]--,t.bl_count[s+1]+=2,t.bl_count[p]--,m-=2}while(0<m);for(s=p;0!==s;s--)for(i=t.bl_count[s];0!==i;)u<(n=t.heap[--r])||(h[2*n+1]!==s&&(t.opt_len+=(s-h[2*n+1])*h[2*n],h[2*n+1]=s),i--)}}(t,e),Z(s,u,t.bl_count)}function X(t,e,r){var i,n,s=-1,a=e[1],o=0,h=7,u=4;for(0===a&&(h=138,u=3),e[2*(r+1)+1]=65535,i=0;i<=r;i++)n=a,a=e[2*(i+1)+1],++o<h&&n===a||(o<u?t.bl_tree[2*n]+=o:0!==n?(n!==s&&t.bl_tree[2*n]++,t.bl_tree[2*b]++):o<=10?t.bl_tree[2*v]++:t.bl_tree[2*y]++,s=n,u=(o=0)===a?(h=138,3):n===a?(h=6,3):(h=7,4))}function V(t,e,r){var i,n,s=-1,a=e[1],o=0,h=7,u=4;for(0===a&&(h=138,u=3),i=0;i<=r;i++)if(n=a,a=e[2*(i+1)+1],!(++o<h&&n===a)){if(o<u)for(;L(t,n,t.bl_tree),0!=--o;);else 0!==n?(n!==s&&(L(t,n,t.bl_tree),o--),L(t,b,t.bl_tree),P(t,o-3,2)):o<=10?(L(t,v,t.bl_tree),P(t,o-3,3)):(L(t,y,t.bl_tree),P(t,o-11,7));s=n,u=(o=0)===a?(h=138,3):n===a?(h=6,3):(h=7,4)}}i(T);var q=!1;function J(t,e,r,i){P(t,(s<<1)+(i?1:0),3),function(t,e,r,i){M(t),i&&(U(t,r),U(t,~r)),n.arraySet(t.pending_buf,t.window,e,r,t.pending),t.pending+=r}(t,e,r,!0)}r._tr_init=function(t){q||(function(){var t,e,r,i,n,s=new Array(g+1);for(i=r=0;i<a-1;i++)for(I[i]=r,t=0;t<1<<w[i];t++)A[r++]=i;for(A[r-1]=i,i=n=0;i<16;i++)for(T[i]=n,t=0;t<1<<k[i];t++)E[n++]=i;for(n>>=7;i<f;i++)for(T[i]=n<<7,t=0;t<1<<k[i]-7;t++)E[256+n++]=i;for(e=0;e<=g;e++)s[e]=0;for(t=0;t<=143;)z[2*t+1]=8,t++,s[8]++;for(;t<=255;)z[2*t+1]=9,t++,s[9]++;for(;t<=279;)z[2*t+1]=7,t++,s[7]++;for(;t<=287;)z[2*t+1]=8,t++,s[8]++;for(Z(z,l+1,s),t=0;t<f;t++)C[2*t+1]=5,C[2*t]=j(t,5);O=new D(z,w,u+1,l,g),B=new D(C,k,0,f,g),R=new D(new Array(0),x,0,d,p)}(),q=!0),t.l_desc=new F(t.dyn_ltree,O),t.d_desc=new F(t.dyn_dtree,B),t.bl_desc=new F(t.bl_tree,R),t.bi_buf=0,t.bi_valid=0,W(t)},r._tr_stored_block=J,r._tr_flush_block=function(t,e,r,i){var n,s,a=0;0<t.level?(2===t.strm.data_type&&(t.strm.data_type=function(t){var e,r=4093624447;for(e=0;e<=31;e++,r>>>=1)if(1&r&&0!==t.dyn_ltree[2*e])return o;if(0!==t.dyn_ltree[18]||0!==t.dyn_ltree[20]||0!==t.dyn_ltree[26])return h;for(e=32;e<u;e++)if(0!==t.dyn_ltree[2*e])return h;return o}(t)),Y(t,t.l_desc),Y(t,t.d_desc),a=function(t){var e;for(X(t,t.dyn_ltree,t.l_desc.max_code),X(t,t.dyn_dtree,t.d_desc.max_code),Y(t,t.bl_desc),e=d-1;3<=e&&0===t.bl_tree[2*S[e]+1];e--);return t.opt_len+=3*(e+1)+5+5+4,e}(t),n=t.opt_len+3+7>>>3,(s=t.static_len+3+7>>>3)<=n&&(n=s)):n=s=r+5,r+4<=n&&-1!==e?J(t,e,r,i):4===t.strategy||s===n?(P(t,2+(i?1:0),3),K(t,z,C)):(P(t,4+(i?1:0),3),function(t,e,r,i){var n;for(P(t,e-257,5),P(t,r-1,5),P(t,i-4,4),n=0;n<i;n++)P(t,t.bl_tree[2*S[n]+1],3);V(t,t.dyn_ltree,e-1),V(t,t.dyn_dtree,r-1)}(t,t.l_desc.max_code+1,t.d_desc.max_code+1,a+1),K(t,t.dyn_ltree,t.dyn_dtree)),W(t),i&&M(t)},r._tr_tally=function(t,e,r){return t.pending_buf[t.d_buf+2*t.last_lit]=e>>>8&255,t.pending_buf[t.d_buf+2*t.last_lit+1]=255&e,t.pending_buf[t.l_buf+t.last_lit]=255&r,t.last_lit++,0===e?t.dyn_ltree[2*r]++:(t.matches++,e--,t.dyn_ltree[2*(A[r]+u+1)]++,t.dyn_dtree[2*N(e)]++),t.last_lit===t.lit_bufsize-1},r._tr_align=function(t){P(t,2,3),L(t,m,z),function(t){16===t.bi_valid?(U(t,t.bi_buf),t.bi_buf=0,t.bi_valid=0):8<=t.bi_valid&&(t.pending_buf[t.pending++]=255&t.bi_buf,t.bi_buf>>=8,t.bi_valid-=8)}(t)}},{"../utils/common":41}],53:[function(t,e,r){"use strict";e.exports=function(){this.input=null,this.next_in=0,this.avail_in=0,this.total_in=0,this.output=null,this.next_out=0,this.avail_out=0,this.total_out=0,this.msg="",this.state=null,this.data_type=2,this.adler=0}},{}],54:[function(t,e,r){"use strict";e.exports="function"==typeof setImmediate?setImmediate:function(){var t=[].slice.apply(arguments);t.splice(1,0,0),setTimeout.apply(null,t)}},{}]},{},[10])(10)});
}).call(this)}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {},require("buffer").Buffer,require("timers").setImmediate)
},{"buffer":2,"timers":8}],7:[function(require,module,exports){
// shim for using process in browser
var process = module.exports = {};

// cached from whatever global is present so that test runners that stub it
// don't break things.  But we need to wrap it in a try catch in case it is
// wrapped in strict mode code which doesn't define any globals.  It's inside a
// function because try/catches deoptimize in certain engines.

var cachedSetTimeout;
var cachedClearTimeout;

function defaultSetTimout() {
    throw new Error('setTimeout has not been defined');
}
function defaultClearTimeout () {
    throw new Error('clearTimeout has not been defined');
}
(function () {
    try {
        if (typeof setTimeout === 'function') {
            cachedSetTimeout = setTimeout;
        } else {
            cachedSetTimeout = defaultSetTimout;
        }
    } catch (e) {
        cachedSetTimeout = defaultSetTimout;
    }
    try {
        if (typeof clearTimeout === 'function') {
            cachedClearTimeout = clearTimeout;
        } else {
            cachedClearTimeout = defaultClearTimeout;
        }
    } catch (e) {
        cachedClearTimeout = defaultClearTimeout;
    }
} ())
function runTimeout(fun) {
    if (cachedSetTimeout === setTimeout) {
        //normal enviroments in sane situations
        return setTimeout(fun, 0);
    }
    // if setTimeout wasn't available but was latter defined
    if ((cachedSetTimeout === defaultSetTimout || !cachedSetTimeout) && setTimeout) {
        cachedSetTimeout = setTimeout;
        return setTimeout(fun, 0);
    }
    try {
        // when when somebody has screwed with setTimeout but no I.E. maddness
        return cachedSetTimeout(fun, 0);
    } catch(e){
        try {
            // When we are in I.E. but the script has been evaled so I.E. doesn't trust the global object when called normally
            return cachedSetTimeout.call(null, fun, 0);
        } catch(e){
            // same as above but when it's a version of I.E. that must have the global object for 'this', hopfully our context correct otherwise it will throw a global error
            return cachedSetTimeout.call(this, fun, 0);
        }
    }


}
function runClearTimeout(marker) {
    if (cachedClearTimeout === clearTimeout) {
        //normal enviroments in sane situations
        return clearTimeout(marker);
    }
    // if clearTimeout wasn't available but was latter defined
    if ((cachedClearTimeout === defaultClearTimeout || !cachedClearTimeout) && clearTimeout) {
        cachedClearTimeout = clearTimeout;
        return clearTimeout(marker);
    }
    try {
        // when when somebody has screwed with setTimeout but no I.E. maddness
        return cachedClearTimeout(marker);
    } catch (e){
        try {
            // When we are in I.E. but the script has been evaled so I.E. doesn't  trust the global object when called normally
            return cachedClearTimeout.call(null, marker);
        } catch (e){
            // same as above but when it's a version of I.E. that must have the global object for 'this', hopfully our context correct otherwise it will throw a global error.
            // Some versions of I.E. have different rules for clearTimeout vs setTimeout
            return cachedClearTimeout.call(this, marker);
        }
    }



}
var queue = [];
var draining = false;
var currentQueue;
var queueIndex = -1;

function cleanUpNextTick() {
    if (!draining || !currentQueue) {
        return;
    }
    draining = false;
    if (currentQueue.length) {
        queue = currentQueue.concat(queue);
    } else {
        queueIndex = -1;
    }
    if (queue.length) {
        drainQueue();
    }
}

function drainQueue() {
    if (draining) {
        return;
    }
    var timeout = runTimeout(cleanUpNextTick);
    draining = true;

    var len = queue.length;
    while(len) {
        currentQueue = queue;
        queue = [];
        while (++queueIndex < len) {
            if (currentQueue) {
                currentQueue[queueIndex].run();
            }
        }
        queueIndex = -1;
        len = queue.length;
    }
    currentQueue = null;
    draining = false;
    runClearTimeout(timeout);
}

process.nextTick = function (fun) {
    var args = new Array(arguments.length - 1);
    if (arguments.length > 1) {
        for (var i = 1; i < arguments.length; i++) {
            args[i - 1] = arguments[i];
        }
    }
    queue.push(new Item(fun, args));
    if (queue.length === 1 && !draining) {
        runTimeout(drainQueue);
    }
};

// v8 likes predictible objects
function Item(fun, array) {
    this.fun = fun;
    this.array = array;
}
Item.prototype.run = function () {
    this.fun.apply(null, this.array);
};
process.title = 'browser';
process.browser = true;
process.env = {};
process.argv = [];
process.version = ''; // empty string to avoid regexp issues
process.versions = {};

function noop() {}

process.on = noop;
process.addListener = noop;
process.once = noop;
process.off = noop;
process.removeListener = noop;
process.removeAllListeners = noop;
process.emit = noop;
process.prependListener = noop;
process.prependOnceListener = noop;

process.listeners = function (name) { return [] }

process.binding = function (name) {
    throw new Error('process.binding is not supported');
};

process.cwd = function () { return '/' };
process.chdir = function (dir) {
    throw new Error('process.chdir is not supported');
};
process.umask = function() { return 0; };

},{}],8:[function(require,module,exports){
(function (setImmediate,clearImmediate){(function (){
var nextTick = require('process/browser.js').nextTick;
var apply = Function.prototype.apply;
var slice = Array.prototype.slice;
var immediateIds = {};
var nextImmediateId = 0;

// DOM APIs, for completeness

exports.setTimeout = function() {
  return new Timeout(apply.call(setTimeout, window, arguments), clearTimeout);
};
exports.setInterval = function() {
  return new Timeout(apply.call(setInterval, window, arguments), clearInterval);
};
exports.clearTimeout =
exports.clearInterval = function(timeout) { timeout.close(); };

function Timeout(id, clearFn) {
  this._id = id;
  this._clearFn = clearFn;
}
Timeout.prototype.unref = Timeout.prototype.ref = function() {};
Timeout.prototype.close = function() {
  this._clearFn.call(window, this._id);
};

// Does not start the time, just sets up the members needed.
exports.enroll = function(item, msecs) {
  clearTimeout(item._idleTimeoutId);
  item._idleTimeout = msecs;
};

exports.unenroll = function(item) {
  clearTimeout(item._idleTimeoutId);
  item._idleTimeout = -1;
};

exports._unrefActive = exports.active = function(item) {
  clearTimeout(item._idleTimeoutId);

  var msecs = item._idleTimeout;
  if (msecs >= 0) {
    item._idleTimeoutId = setTimeout(function onTimeout() {
      if (item._onTimeout)
        item._onTimeout();
    }, msecs);
  }
};

// That's not how node.js implements it but the exposed api is the same.
exports.setImmediate = typeof setImmediate === "function" ? setImmediate : function(fn) {
  var id = nextImmediateId++;
  var args = arguments.length < 2 ? false : slice.call(arguments, 1);

  immediateIds[id] = true;

  nextTick(function onNextTick() {
    if (immediateIds[id]) {
      // fn.call() is faster so we optimize for the common use-case
      // @see http://jsperf.com/call-apply-segu
      if (args) {
        fn.apply(null, args);
      } else {
        fn.call(null);
      }
      // Prevent ids from leaking
      exports.clearImmediate(id);
    }
  });

  return id;
};

exports.clearImmediate = typeof clearImmediate === "function" ? clearImmediate : function(id) {
  delete immediateIds[id];
};
}).call(this)}).call(this,require("timers").setImmediate,require("timers").clearImmediate)
},{"process/browser.js":7,"timers":8}],9:[function(require,module,exports){
"use strict";

var __importDefault = void 0 && (void 0).__importDefault || function (mod) {
  return mod && mod.__esModule ? mod : {
    "default": mod
  };
};

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.HandleZip = void 0;

var jszip_1 = __importDefault(require("jszip"));

var method_1 = require("./common/method");

var HandleZip =
/** @class */
function () {
  function HandleZip(file) {
    // Support nodejs fs to read files
    // if(file instanceof File){
    this.uploadFile = file; // }
  }

  HandleZip.prototype.unzipFile = function (successFunc, errorFunc) {
    // var new_zip:JSZip = new JSZip();
    jszip_1["default"].loadAsync(this.uploadFile) // 1) read the Blob
    .then(function (zip) {
      var fileList = {},
          lastIndex = Object.keys(zip.files).length,
          index = 0;
      zip.forEach(function (relativePath, zipEntry) {
        var fileName = zipEntry.name;
        var fileNameArr = fileName.split(".");
        var suffix = fileNameArr[fileNameArr.length - 1].toLowerCase();
        var fileType = "string";

        if (suffix in {
          "png": 1,
          "jpeg": 1,
          "jpg": 1,
          "gif": 1,
          "bmp": 1,
          "tif": 1,
          "webp": 1
        }) {
          fileType = "base64";
        } else if (suffix == "emf") {
          fileType = "arraybuffer";
        }

        zipEntry.async(fileType).then(function (data) {
          if (fileType == "base64") {
            data = "data:image/" + suffix + ";base64," + data;
          }

          fileList[zipEntry.name] = data; // console.log(lastIndex, index);

          if (lastIndex == index + 1) {
            successFunc(fileList);
          }

          index++;
        });
      });
    }, function (e) {
      errorFunc(e);
    });
  };

  HandleZip.prototype.unzipFileByUrl = function (url, successFunc, errorFunc) {
    var new_zip = new jszip_1["default"]();
    method_1.getBinaryContent(url, function (err, data) {
      if (err) {
        throw err; // or handle err
      }

      jszip_1["default"].loadAsync(data).then(function (zip) {
        var fileList = {},
            lastIndex = Object.keys(zip.files).length,
            index = 0;
        zip.forEach(function (relativePath, zipEntry) {
          var fileName = zipEntry.name;
          var fileNameArr = fileName.split(".");
          var suffix = fileNameArr[fileNameArr.length - 1].toLowerCase();
          var fileType = "string";

          if (suffix in {
            "png": 1,
            "jpeg": 1,
            "jpg": 1,
            "gif": 1,
            "bmp": 1,
            "tif": 1,
            "webp": 1
          }) {
            fileType = "base64";
          } else if (suffix == "emf") {
            fileType = "arraybuffer";
          }

          zipEntry.async(fileType).then(function (data) {
            if (fileType == "base64") {
              data = "data:image/" + suffix + ";base64," + data;
            }

            fileList[zipEntry.name] = data; // console.log(lastIndex, index);

            if (lastIndex == index + 1) {
              successFunc(fileList);
            }

            index++;
          });
        });
      }, function (e) {
        errorFunc(e);
      });
    });
  };

  HandleZip.prototype.newZipFile = function () {
    var zip = new jszip_1["default"]();
    this.workBook = zip;
  }; //title:"nested/hello.txt", content:"Hello Worldasdfasfasdfasfasfasfasfasdfas"


  HandleZip.prototype.addToZipFile = function (title, content) {
    if (this.workBook == null) {
      var zip = new jszip_1["default"]();
      this.workBook = zip;
    }

    this.workBook.file(title, content);
  };

  return HandleZip;
}();

exports.HandleZip = HandleZip;

},{"./common/method":18,"jszip":6}],10:[function(require,module,exports){
"use strict";

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.LuckyImageBase = exports.LuckysheetCalcChain = exports.LuckySheetConfigMerge = exports.LuckySheetborderInfoCellValueStyle = exports.LuckySheetborderInfoCellValue = exports.LuckySheetborderInfoCellForImp = exports.LuckyConfig = exports.LuckyInlineString = exports.LuckySheetCellFormat = exports.LuckySheetCelldataValue = exports.LuckySheetCelldataBase = exports.LuckyFileInfo = exports.LuckySheetBase = exports.LuckyFileBase = void 0;

var LuckyFileBase =
/** @class */
function () {
  function LuckyFileBase() {}

  return LuckyFileBase;
}();

exports.LuckyFileBase = LuckyFileBase;

var LuckySheetBase =
/** @class */
function () {
  function LuckySheetBase() {}

  return LuckySheetBase;
}();

exports.LuckySheetBase = LuckySheetBase;

var LuckyFileInfo =
/** @class */
function () {
  function LuckyFileInfo() {}

  return LuckyFileInfo;
}();

exports.LuckyFileInfo = LuckyFileInfo;

var LuckySheetCelldataBase =
/** @class */
function () {
  function LuckySheetCelldataBase() {}

  return LuckySheetCelldataBase;
}();

exports.LuckySheetCelldataBase = LuckySheetCelldataBase;

var LuckySheetCelldataValue =
/** @class */
function () {
  function LuckySheetCelldataValue() {}

  return LuckySheetCelldataValue;
}();

exports.LuckySheetCelldataValue = LuckySheetCelldataValue;

var LuckySheetCellFormat =
/** @class */
function () {
  function LuckySheetCellFormat() {}

  return LuckySheetCellFormat;
}();

exports.LuckySheetCellFormat = LuckySheetCellFormat;

var LuckyInlineString =
/** @class */
function () {
  function LuckyInlineString() {}

  return LuckyInlineString;
}();

exports.LuckyInlineString = LuckyInlineString;

var LuckyConfig =
/** @class */
function () {
  function LuckyConfig() {}

  return LuckyConfig;
}();

exports.LuckyConfig = LuckyConfig;

var LuckySheetborderInfoCellForImp =
/** @class */
function () {
  function LuckySheetborderInfoCellForImp() {}

  return LuckySheetborderInfoCellForImp;
}();

exports.LuckySheetborderInfoCellForImp = LuckySheetborderInfoCellForImp;

var LuckySheetborderInfoCellValue =
/** @class */
function () {
  function LuckySheetborderInfoCellValue() {}

  return LuckySheetborderInfoCellValue;
}();

exports.LuckySheetborderInfoCellValue = LuckySheetborderInfoCellValue;

var LuckySheetborderInfoCellValueStyle =
/** @class */
function () {
  function LuckySheetborderInfoCellValueStyle() {}

  return LuckySheetborderInfoCellValueStyle;
}();

exports.LuckySheetborderInfoCellValueStyle = LuckySheetborderInfoCellValueStyle;

var LuckySheetConfigMerge =
/** @class */
function () {
  function LuckySheetConfigMerge() {}

  return LuckySheetConfigMerge;
}();

exports.LuckySheetConfigMerge = LuckySheetConfigMerge;

var LuckysheetCalcChain =
/** @class */
function () {
  function LuckysheetCalcChain() {}

  return LuckysheetCalcChain;
}();

exports.LuckysheetCalcChain = LuckysheetCalcChain;

var LuckyImageBase =
/** @class */
function () {
  function LuckyImageBase() {}

  return LuckyImageBase;
}();

exports.LuckyImageBase = LuckyImageBase;

},{}],11:[function(require,module,exports){
"use strict";

var __extends = void 0 && (void 0).__extends || function () {
  var _extendStatics = function extendStatics(d, b) {
    _extendStatics = Object.setPrototypeOf || {
      __proto__: []
    } instanceof Array && function (d, b) {
      d.__proto__ = b;
    } || function (d, b) {
      for (var p in b) {
        if (b.hasOwnProperty(p)) d[p] = b[p];
      }
    };

    return _extendStatics(d, b);
  };

  return function (d, b) {
    _extendStatics(d, b);

    function __() {
      this.constructor = d;
    }

    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
  };
}();

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.LuckySheetCelldata = void 0;

var ReadXml_1 = require("./ReadXml");

var method_1 = require("../common/method");

var constant_1 = require("../common/constant");

var LuckyBase_1 = require("./LuckyBase");

var LuckySheetCelldata =
/** @class */
function (_super) {
  __extends(LuckySheetCelldata, _super);

  function LuckySheetCelldata(cell, styles, sharedStrings, mergeCells, sheetFile, ReadXml) {
    var _this = //Private
    _super.call(this) || this;

    _this.cell = cell;
    _this.sheetFile = sheetFile;
    _this.styles = styles;
    _this.sharedStrings = sharedStrings;
    _this.readXml = ReadXml;
    _this.mergeCells = mergeCells;
    var attrList = cell.attributeList;
    var r = attrList.r,
        s = attrList.s,
        t = attrList.t;
    var range = method_1.getcellrange(r);
    _this.r = range.row[0];
    _this.c = range.column[0];
    _this.v = _this.generateValue(s, t);
    return _this;
  }
  /**
  * @param s Style index ,start 1
  * @param t Cell type, Optional value is ST_CellType, it's found at constat.ts
  */


  LuckySheetCelldata.prototype.generateValue = function (s, t) {
    var _this = this;

    var v = this.cell.getInnerElements("v");
    var f = this.cell.getInnerElements("f");

    if (v == null) {
      v = this.cell.getInnerElements("t");
    }

    var cellXfs = this.styles["cellXfs"];
    var cellStyleXfs = this.styles["cellStyleXfs"];
    var cellStyles = this.styles["cellStyles"];
    var fonts = this.styles["fonts"];
    var fills = this.styles["fills"];
    var borders = this.styles["borders"];
    var numfmts = this.styles["numfmts"];
    var clrScheme = this.styles["clrScheme"];
    var sharedStrings = this.sharedStrings;
    var cellValue = new LuckyBase_1.LuckySheetCelldataValue();

    if (f != null) {
      var formula = f[0],
          attrList = formula.attributeList;
      var t_1 = attrList.t,
          ref = attrList.ref,
          si = attrList.si;
      var formulaValue = f[0].value;

      if (t_1 == "shared") {
        this._fomulaRef = ref;
        this._formulaType = t_1;
        this._formulaSi = si;
      } // console.log(ref, t, si);


      if (ref != null || formulaValue != null && formulaValue.length > 0) {
        formulaValue = method_1.escapeCharacter(formulaValue);
        cellValue.f = "=" + formulaValue;
      }
    }

    var familyFont = null;
    var quotePrefix;

    if (s != null) {
      var sNum = parseInt(s);
      var cellXf = cellXfs[sNum];
      var xfId = cellXf.attributeList.xfId;
      var numFmtId = void 0,
          fontId = void 0,
          fillId = void 0,
          borderId = void 0;
      var horizontal = void 0,
          vertical = void 0,
          wrapText = void 0,
          textRotation = void 0,
          shrinkToFit = void 0,
          indent = void 0,
          applyProtection = void 0;

      if (xfId != null) {
        var cellStyleXf = cellStyleXfs[parseInt(xfId)];
        var attrList = cellStyleXf.attributeList;
        var applyNumberFormat_1 = attrList.applyNumberFormat;
        var applyFont_1 = attrList.applyFont;
        var applyFill_1 = attrList.applyFill;
        var applyBorder_1 = attrList.applyBorder;
        var applyAlignment_1 = attrList.applyAlignment; // let applyProtection = attrList.applyProtection;

        applyProtection = attrList.applyProtection;
        quotePrefix = attrList.quotePrefix;

        if (applyNumberFormat_1 != "0" && attrList.numFmtId != null) {
          // if(attrList.numFmtId!="0"){
          numFmtId = attrList.numFmtId; // }
        }

        if (applyFont_1 != "0" && attrList.fontId != null) {
          fontId = attrList.fontId;
        }

        if (applyFill_1 != "0" && attrList.fillId != null) {
          fillId = attrList.fillId;
        }

        if (applyBorder_1 != "0" && attrList.borderId != null) {
          borderId = attrList.borderId;
        }

        if (applyAlignment_1 != null && applyAlignment_1 != "0") {
          var alignment = cellStyleXf.getInnerElements("alignment");

          if (alignment != null) {
            var attrList_1 = alignment[0].attributeList;

            if (attrList_1.horizontal != null) {
              horizontal = attrList_1.horizontal;
            }

            if (attrList_1.vertical != null) {
              vertical = attrList_1.vertical;
            }

            if (attrList_1.wrapText != null) {
              wrapText = attrList_1.wrapText;
            }

            if (attrList_1.textRotation != null) {
              textRotation = attrList_1.textRotation;
            }

            if (attrList_1.shrinkToFit != null) {
              shrinkToFit = attrList_1.shrinkToFit;
            }

            if (attrList_1.indent != null) {
              indent = attrList_1.indent;
            }
          }
        }
      }

      var applyNumberFormat = cellXf.attributeList.applyNumberFormat;
      var applyFont = cellXf.attributeList.applyFont;
      var applyFill = cellXf.attributeList.applyFill;
      var applyBorder = cellXf.attributeList.applyBorder;
      var applyAlignment = cellXf.attributeList.applyAlignment;

      if (cellXf.attributeList.applyProtection != null) {
        applyProtection = cellXf.attributeList.applyProtection;
      }

      if (cellXf.attributeList.quotePrefix != null) {
        quotePrefix = cellXf.attributeList.quotePrefix;
      }

      if (applyNumberFormat != "0" && cellXf.attributeList.numFmtId != null) {
        numFmtId = cellXf.attributeList.numFmtId;
      }

      if (applyFont != "0") {
        fontId = cellXf.attributeList.fontId;
      }

      if (applyFill != "0") {
        fillId = cellXf.attributeList.fillId;
      }

      if (applyBorder != "0") {
        borderId = cellXf.attributeList.borderId;
      }

      if (applyAlignment != "0") {
        var alignment = cellXf.getInnerElements("alignment");

        if (alignment != null && alignment.length > 0) {
          var attrList = alignment[0].attributeList;

          if (attrList.horizontal != null) {
            horizontal = attrList.horizontal;
          }

          if (attrList.vertical != null) {
            vertical = attrList.vertical;
          }

          if (attrList.wrapText != null) {
            wrapText = attrList.wrapText;
          }

          if (attrList.textRotation != null) {
            textRotation = attrList.textRotation;
          }

          if (attrList.shrinkToFit != null) {
            shrinkToFit = attrList.shrinkToFit;
          }

          if (attrList.indent != null) {
            indent = attrList.indent;
          }
        }
      }

      if (numFmtId != undefined) {
        var numf = numfmts[parseInt(numFmtId)];
        var cellFormat = new LuckyBase_1.LuckySheetCellFormat();
        cellFormat.fa = method_1.escapeCharacter(numf); // console.log(numf, numFmtId, this.v);

        cellFormat.t = t || 'd';
        cellValue.ct = cellFormat;
      }

      if (fillId != undefined) {
        var fillIdNum = parseInt(fillId);
        var fill = fills[fillIdNum]; // console.log(cellValue.v);

        var bg = this.getBackgroundByFill(fill, clrScheme);

        if (bg != null) {
          cellValue.bg = bg;
        }
      }

      if (fontId != undefined) {
        var fontIdNum = parseInt(fontId);
        var font = fonts[fontIdNum];

        if (font != null) {
          var sz = font.getInnerElements("sz"); //font size

          var colors = font.getInnerElements("color"); //font color

          var family = font.getInnerElements("name"); //font family

          var familyOverrides = font.getInnerElements("family"); //font family will be overrided by name

          var charset = font.getInnerElements("charset"); //font charset

          var bolds = font.getInnerElements("b"); //font bold

          var italics = font.getInnerElements("i"); //font italic

          var strikes = font.getInnerElements("strike"); //font italic

          var underlines = font.getInnerElements("u"); //font italic

          if (sz != null && sz.length > 0) {
            var fs = sz[0].attributeList.val;

            if (fs != null) {
              cellValue.fs = parseInt(fs);
            }
          }

          if (colors != null && colors.length > 0) {
            var color = colors[0];
            var fc = ReadXml_1.getColor(color, this.styles, "t");

            if (fc != null) {
              cellValue.fc = fc;
            }
          }

          if (familyOverrides != null && familyOverrides.length > 0) {
            var val = familyOverrides[0].attributeList.val;

            if (val != null) {
              familyFont = constant_1.fontFamilys[val];
            }
          }

          if (family != null && family.length > 0) {
            var val = family[0].attributeList.val;

            if (val != null) {
              cellValue.ff = val;
            }
          }

          if (bolds != null && bolds.length > 0) {
            var bold = bolds[0].attributeList.val;

            if (bold == "0") {
              cellValue.bl = 0;
            } else {
              cellValue.bl = 1;
            }
          }

          if (italics != null && italics.length > 0) {
            var italic = italics[0].attributeList.val;

            if (italic == "0") {
              cellValue.it = 0;
            } else {
              cellValue.it = 1;
            }
          }

          if (strikes != null && strikes.length > 0) {
            var strike = strikes[0].attributeList.val;

            if (strike == "0") {
              cellValue.cl = 0;
            } else {
              cellValue.cl = 1;
            }
          }

          if (underlines != null && underlines.length > 0) {
            var underline = underlines[0].attributeList.val;

            if (underline == "single") {
              cellValue.un = 1;
            } else if (underline == "double") {
              cellValue.un = 2;
            } else if (underline == "singleAccounting") {
              cellValue.un = 3;
            } else if (underline == "doubleAccounting") {
              cellValue.un = 4;
            } else {
              cellValue.un = 0;
            }
          }
        }
      } // vt: number | undefined//Vertical alignment, 0 middle, 1 up, 2 down, alignment
      // ht: number | undefined//Horizontal alignment,0 center, 1 left, 2 right, alignment
      // tr: number | undefined //Text rotation,0: 0、1: 45 、2: -45、3 Vertical text、4: 90 、5: -90, alignment
      // tb: number | undefined //Text wrap,0 truncation, 1 overflow, 2 word wrap, alignment


      if (horizontal != undefined) {
        //Horizontal alignment
        if (horizontal == "center") {
          cellValue.ht = 0;
        } else if (horizontal == "centerContinuous") {
          cellValue.ht = 0; //luckysheet unsupport
        } else if (horizontal == "left") {
          cellValue.ht = 1;
        } else if (horizontal == "right") {
          cellValue.ht = 2;
        } else if (horizontal == "distributed") {
          cellValue.ht = 0; //luckysheet unsupport
        } else if (horizontal == "fill") {
          cellValue.ht = 1; //luckysheet unsupport
        } else if (horizontal == "general") {
          cellValue.ht = 1; //luckysheet unsupport
        } else if (horizontal == "justify") {
          cellValue.ht = 0; //luckysheet unsupport
        } else {
          cellValue.ht = 1;
        }
      }

      if (vertical != undefined) {
        //Vertical alignment
        if (vertical == "bottom") {
          cellValue.vt = 2;
        } else if (vertical == "center") {
          cellValue.vt = 0;
        } else if (vertical == "distributed") {
          cellValue.vt = 0; //luckysheet unsupport
        } else if (vertical == "justify") {
          cellValue.vt = 0; //luckysheet unsupport
        } else if (vertical == "top") {
          cellValue.vt = 1;
        } else {
          cellValue.vt = 1;
        }
      } else {
        //sometimes bottom style is lost after setting it in excel
        //when vertical is undefined set it to 2.
        cellValue.vt = 2;
      }

      if (wrapText != undefined) {
        if (wrapText == "1") {
          cellValue.tb = 2;
        } else {
          cellValue.tb = 1;
        }
      } else {
        cellValue.tb = 1;
      }

      if (textRotation != undefined) {
        // tr: number | undefined //Text rotation,0: 0、1: 45 、2: -45、3 Vertical text、4: 90 、5: -90, alignment
        if (textRotation == "255") {
          cellValue.tr = 3;
        } // else if(textRotation=="45"){
        //     cellValue.tr = 1;
        // }
        // else if(textRotation=="90"){
        //     cellValue.tr = 4;
        // }
        // else if(textRotation=="135"){
        //     cellValue.tr = 2;
        // }
        // else if(textRotation=="180"){
        //     cellValue.tr = 5;
        // }
        else {
          cellValue.tr = 0;
          cellValue.rt = parseInt(textRotation);
        }
      }

      if (shrinkToFit != undefined) {//luckysheet unsupport
      }

      if (indent != undefined) {//luckysheet unsupport
      }

      if (borderId != undefined) {
        var borderIdNum = parseInt(borderId);
        var border = borders[borderIdNum]; // this._borderId = borderIdNum;

        var borderObject = new LuckyBase_1.LuckySheetborderInfoCellForImp();
        borderObject.rangeType = "cell"; // borderObject.cells = [];

        var borderCellValue = new LuckyBase_1.LuckySheetborderInfoCellValue();
        borderCellValue.row_index = this.r;
        borderCellValue.col_index = this.c;
        var lefts = border.getInnerElements("left");
        var rights = border.getInnerElements("right");
        var tops = border.getInnerElements("top");
        var bottoms = border.getInnerElements("bottom");
        var diagonals = border.getInnerElements("diagonal");
        var starts = border.getInnerElements("start");
        var ends = border.getInnerElements("end");
        var left = this.getBorderInfo(lefts);
        var right = this.getBorderInfo(rights);
        var top_1 = this.getBorderInfo(tops);
        var bottom = this.getBorderInfo(bottoms);
        var diagonal = this.getBorderInfo(diagonals);
        var start = this.getBorderInfo(starts);
        var end = this.getBorderInfo(ends);
        var isAdd = false;

        if (start != null && start.color != null) {
          borderCellValue.l = start;
          isAdd = true;
        }

        if (end != null && end.color != null) {
          borderCellValue.r = end;
          isAdd = true;
        }

        if (left != null && left.color != null) {
          borderCellValue.l = left;
          isAdd = true;
        }

        if (right != null && right.color != null) {
          borderCellValue.r = right;
          isAdd = true;
        }

        if (top_1 != null && top_1.color != null) {
          borderCellValue.t = top_1;
          isAdd = true;
        }

        if (bottom != null && bottom.color != null) {
          borderCellValue.b = bottom;
          isAdd = true;
        }

        if (isAdd) {
          borderObject.value = borderCellValue; // this.config._borderInfo[borderId] = borderObject;

          this._borderObject = borderObject;
        }
      }
    } else {
      cellValue.tb = 1;
    }

    if (v != null) {
      var value = v[0].value;

      if (/&#\d+;/.test(value)) {
        value = this.htmlDecode(value);
      }

      if (t == constant_1.ST_CellType["SharedString"]) {
        var siIndex = parseInt(v[0].value);
        var sharedSI = sharedStrings[siIndex];
        var rFlag = sharedSI.getInnerElements("r");

        if (rFlag == null) {
          var tFlag = sharedSI.getInnerElements("t");

          if (tFlag != null) {
            var text_1 = "";
            tFlag.forEach(function (t) {
              text_1 += t.value;
            });
            text_1 = method_1.escapeCharacter(text_1); //isContainMultiType(text) &&

            if (familyFont == "Roman" && text_1.length > 0) {
              var textArray = text_1.split("");
              var preWordType = null,
                  wordText = "",
                  preWholef = null;
              var wholef = "Times New Roman";

              if (cellValue.ff != null) {
                wholef = cellValue.ff;
              }

              var cellFormat = cellValue.ct;

              if (cellFormat == null) {
                cellFormat = new LuckyBase_1.LuckySheetCellFormat();
              }

              if (cellFormat.s == null) {
                cellFormat.s = [];
              }

              for (var i = 0; i < textArray.length; i++) {
                var w = textArray[i];
                var type = null,
                    ff = wholef;

                if (method_1.isChinese(w)) {
                  type = "c";
                  ff = "宋体";
                } else if (method_1.isJapanese(w)) {
                  type = "j";
                  ff = "Yu Gothic";
                } else if (method_1.isKoera(w)) {
                  type = "k";
                  ff = "Malgun Gothic";
                } else {
                  type = "e";
                }

                if (type != preWordType && preWordType != null || i == textArray.length - 1) {
                  var InlineString = new LuckyBase_1.LuckyInlineString();
                  InlineString.ff = preWholef;

                  if (cellValue.fc != null) {
                    InlineString.fc = cellValue.fc;
                  }

                  if (cellValue.fs != null) {
                    InlineString.fs = cellValue.fs;
                  }

                  if (cellValue.cl != null) {
                    InlineString.cl = cellValue.cl;
                  }

                  if (cellValue.un != null) {
                    InlineString.un = cellValue.un;
                  }

                  if (cellValue.bl != null) {
                    InlineString.bl = cellValue.bl;
                  }

                  if (cellValue.it != null) {
                    InlineString.it = cellValue.it;
                  }

                  if (i == textArray.length - 1) {
                    if (type == preWordType) {
                      InlineString.ff = ff;
                      InlineString.v = wordText + w;
                    } else {
                      InlineString.ff = preWholef;
                      InlineString.v = wordText;
                      cellFormat.s.push(InlineString);
                      var InlineStringLast = new LuckyBase_1.LuckyInlineString();
                      InlineStringLast.ff = ff;
                      InlineStringLast.v = w;

                      if (cellValue.fc != null) {
                        InlineStringLast.fc = cellValue.fc;
                      }

                      if (cellValue.fs != null) {
                        InlineStringLast.fs = cellValue.fs;
                      }

                      if (cellValue.cl != null) {
                        InlineStringLast.cl = cellValue.cl;
                      }

                      if (cellValue.un != null) {
                        InlineStringLast.un = cellValue.un;
                      }

                      if (cellValue.bl != null) {
                        InlineStringLast.bl = cellValue.bl;
                      }

                      if (cellValue.it != null) {
                        InlineStringLast.it = cellValue.it;
                      }

                      cellFormat.s.push(InlineStringLast);
                      break;
                    }
                  } else {
                    InlineString.v = wordText;
                  }

                  cellFormat.s.push(InlineString);
                  wordText = w;
                } else {
                  wordText += w;
                }

                preWordType = type;
                preWholef = ff;
              }

              cellFormat.t = "inlineStr"; // cellFormat.s = [InlineString];

              cellValue.ct = cellFormat; // console.log(cellValue);
            } else {
              text_1 = this.replaceSpecialWrap(text_1);

              if (text_1.indexOf("\r\n") > -1 || text_1.indexOf("\n") > -1) {
                var InlineString = new LuckyBase_1.LuckyInlineString();
                InlineString.v = text_1;
                var cellFormat = cellValue.ct;

                if (cellFormat == null) {
                  cellFormat = new LuckyBase_1.LuckySheetCellFormat();
                }

                if (cellValue.ff != null) {
                  InlineString.ff = cellValue.ff;
                }

                if (cellValue.fc != null) {
                  InlineString.fc = cellValue.fc;
                }

                if (cellValue.fs != null) {
                  InlineString.fs = cellValue.fs;
                }

                if (cellValue.cl != null) {
                  InlineString.cl = cellValue.cl;
                }

                if (cellValue.un != null) {
                  InlineString.un = cellValue.un;
                }

                if (cellValue.bl != null) {
                  InlineString.bl = cellValue.bl;
                }

                if (cellValue.it != null) {
                  InlineString.it = cellValue.it;
                }

                cellFormat.t = "inlineStr";
                cellFormat.s = [InlineString];
                cellValue.ct = cellFormat;
              } else {
                cellValue.v = text_1;
                quotePrefix = "1";
              }
            }
          }
        } else {
          var styles_1 = [];
          rFlag.forEach(function (r) {
            var tFlag = r.getInnerElements("t");
            var rPr = r.getInnerElements("rPr");
            var InlineString = new LuckyBase_1.LuckyInlineString();

            if (tFlag != null && tFlag.length > 0) {
              var text = tFlag[0].value;
              text = _this.replaceSpecialWrap(text);
              text = method_1.escapeCharacter(text);
              InlineString.v = text;
            }

            if (rPr != null && rPr.length > 0) {
              var frpr = rPr[0];
              var sz = ReadXml_1.getlineStringAttr(frpr, "sz"),
                  rFont = ReadXml_1.getlineStringAttr(frpr, "rFont"),
                  family = ReadXml_1.getlineStringAttr(frpr, "family"),
                  charset = ReadXml_1.getlineStringAttr(frpr, "charset"),
                  scheme = ReadXml_1.getlineStringAttr(frpr, "scheme"),
                  b = ReadXml_1.getlineStringAttr(frpr, "b"),
                  i = ReadXml_1.getlineStringAttr(frpr, "i"),
                  u = ReadXml_1.getlineStringAttr(frpr, "u"),
                  strike = ReadXml_1.getlineStringAttr(frpr, "strike"),
                  vertAlign = ReadXml_1.getlineStringAttr(frpr, "vertAlign"),
                  color = void 0;
              var cEle = frpr.getInnerElements("color");

              if (cEle != null && cEle.length > 0) {
                color = ReadXml_1.getColor(cEle[0], _this.styles, "t");
              }

              var ff = void 0; // if(family!=null){
              //     ff = fontFamilys[family];
              // }

              if (rFont != null) {
                ff = rFont;
              }

              if (ff != null) {
                InlineString.ff = ff;
              } else if (cellValue.ff != null) {
                InlineString.ff = cellValue.ff;
              }

              if (color != null) {
                InlineString.fc = color;
              } else if (cellValue.fc != null) {
                InlineString.fc = cellValue.fc;
              }

              if (sz != null) {
                InlineString.fs = parseInt(sz);
              } else if (cellValue.fs != null) {
                InlineString.fs = cellValue.fs;
              }

              if (strike != null) {
                InlineString.cl = parseInt(strike);
              } else if (cellValue.cl != null) {
                InlineString.cl = cellValue.cl;
              }

              if (u != null) {
                InlineString.un = parseInt(u);
              } else if (cellValue.un != null) {
                InlineString.un = cellValue.un;
              }

              if (b != null) {
                InlineString.bl = parseInt(b);
              } else if (cellValue.bl != null) {
                InlineString.bl = cellValue.bl;
              }

              if (i != null) {
                InlineString.it = parseInt(i);
              } else if (cellValue.it != null) {
                InlineString.it = cellValue.it;
              }

              if (vertAlign != null) {
                InlineString.va = parseInt(vertAlign);
              } // ff:string | undefined //font family
              // fc:string | undefined//font color
              // fs:number | undefined//font size
              // cl:number | undefined//strike
              // un:number | undefined//underline
              // bl:number | undefined//blod
              // it:number | undefined//italic
              // v:string | undefined

            } else {
              if (InlineString.ff == null && cellValue.ff != null) {
                InlineString.ff = cellValue.ff;
              }

              if (InlineString.fc == null && cellValue.fc != null) {
                InlineString.fc = cellValue.fc;
              }

              if (InlineString.fs == null && cellValue.fs != null) {
                InlineString.fs = cellValue.fs;
              }

              if (InlineString.cl == null && cellValue.cl != null) {
                InlineString.cl = cellValue.cl;
              }

              if (InlineString.un == null && cellValue.un != null) {
                InlineString.un = cellValue.un;
              }

              if (InlineString.bl == null && cellValue.bl != null) {
                InlineString.bl = cellValue.bl;
              }

              if (InlineString.it == null && cellValue.it != null) {
                InlineString.it = cellValue.it;
              }
            }

            styles_1.push(InlineString);
          });
          var cellFormat = cellValue.ct;

          if (cellFormat == null) {
            cellFormat = new LuckyBase_1.LuckySheetCellFormat();
          }

          cellFormat.t = "inlineStr";
          cellFormat.s = styles_1;
          cellValue.ct = cellFormat;
        }
      } // to be confirmed
      else if (t == constant_1.ST_CellType["InlineString"] && v != null) {
        cellValue.v = "'" + value;
      } else {
        value = method_1.escapeCharacter(value);
        cellValue.v = value;
      }
    }

    if (quotePrefix != null) {
      cellValue.qp = parseInt(quotePrefix);
    }

    return cellValue;
  };

  LuckySheetCelldata.prototype.replaceSpecialWrap = function (text) {
    text = text.replace(/_x000D_/g, "").replace(/&#13;&#10;/g, "\r\n").replace(/&#13;/g, "\r").replace(/&#10;/g, "\n");
    return text;
  };

  LuckySheetCelldata.prototype.getBackgroundByFill = function (fill, clrScheme) {
    var patternFills = fill.getInnerElements("patternFill");

    if (patternFills != null) {
      var patternFill = patternFills[0];
      var fgColors = patternFill.getInnerElements("fgColor");
      var bgColors = patternFill.getInnerElements("bgColor");
      var fg = void 0,
          bg = void 0;

      if (fgColors != null) {
        var fgColor = fgColors[0];
        fg = ReadXml_1.getColor(fgColor, this.styles);
      }

      if (bgColors != null) {
        var bgColor = bgColors[0];
        bg = ReadXml_1.getColor(bgColor, this.styles);
      } // console.log(fgColors,bgColors,clrScheme);


      if (fg != null) {
        return fg;
      } else if (bg != null) {
        return bg;
      }
    } else {
      var gradientfills = fill.getInnerElements("gradientFill");

      if (gradientfills != null) {
        //graient color fill handler
        return null;
      }
    }
  };

  LuckySheetCelldata.prototype.getBorderInfo = function (borders) {
    if (borders == null) {
      return null;
    }

    var border = borders[0],
        attrList = border.attributeList;
    var clrScheme = this.styles["clrScheme"];
    var style = attrList.style;

    if (style == null || style == "none") {
      return null;
    }

    var colors = border.getInnerElements("color");
    var colorRet = "#000000";

    if (colors != null) {
      var color = colors[0];
      colorRet = ReadXml_1.getColor(color, this.styles, "b");

      if (colorRet == null) {
        colorRet = "#000000";
      }
    }

    var ret = new LuckyBase_1.LuckySheetborderInfoCellValueStyle();
    ret.style = constant_1.borderTypes[style];
    ret.color = colorRet;
    return ret;
  };

  LuckySheetCelldata.prototype.htmlDecode = function (str) {
    return str.replace(/&#(x)?([^&]{1,5});/g, function ($, $1, $2) {
      return String.fromCharCode(parseInt($2, $1 ? 16 : 10));
    });
  };

  ;
  return LuckySheetCelldata;
}(LuckyBase_1.LuckySheetCelldataBase);

exports.LuckySheetCelldata = LuckySheetCelldata;

},{"../common/constant":16,"../common/method":18,"./LuckyBase":10,"./ReadXml":15}],12:[function(require,module,exports){
"use strict";

var __extends = void 0 && (void 0).__extends || function () {
  var _extendStatics = function extendStatics(d, b) {
    _extendStatics = Object.setPrototypeOf || {
      __proto__: []
    } instanceof Array && function (d, b) {
      d.__proto__ = b;
    } || function (d, b) {
      for (var p in b) {
        if (b.hasOwnProperty(p)) d[p] = b[p];
      }
    };

    return _extendStatics(d, b);
  };

  return function (d, b) {
    _extendStatics(d, b);

    function __() {
      this.constructor = d;
    }

    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
  };
}();

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.LuckyFile = void 0;

var LuckySheet_1 = require("./LuckySheet");

var constant_1 = require("../common/constant");

var ReadXml_1 = require("./ReadXml");

var method_1 = require("../common/method");

var LuckyBase_1 = require("./LuckyBase");

var LuckyImage_1 = require("./LuckyImage");

var LuckyFile =
/** @class */
function (_super) {
  __extends(LuckyFile, _super);

  function LuckyFile(files, fileName) {
    var _this = _super.call(this) || this;

    _this.columnWidthSet = [];
    _this.rowHeightSet = [];
    _this.files = files;
    _this.fileName = fileName;
    _this.readXml = new ReadXml_1.ReadXml(files);

    _this.getSheetNameList();

    _this.sharedStrings = _this.readXml.getElementsByTagName("sst/si", constant_1.sharedStringsFile);
    _this.calcChain = _this.readXml.getElementsByTagName("calcChain/c", constant_1.calcChainFile);
    _this.styles = {};
    _this.styles["cellXfs"] = _this.readXml.getElementsByTagName("cellXfs/xf", constant_1.stylesFile);
    _this.styles["cellStyleXfs"] = _this.readXml.getElementsByTagName("cellStyleXfs/xf", constant_1.stylesFile);
    _this.styles["cellStyles"] = _this.readXml.getElementsByTagName("cellStyles/cellStyle", constant_1.stylesFile);
    _this.styles["fonts"] = _this.readXml.getElementsByTagName("fonts/font", constant_1.stylesFile);
    _this.styles["fills"] = _this.readXml.getElementsByTagName("fills/fill", constant_1.stylesFile);
    _this.styles["borders"] = _this.readXml.getElementsByTagName("borders/border", constant_1.stylesFile);
    _this.styles["clrScheme"] = _this.readXml.getElementsByTagName("a:clrScheme/a:dk1|a:lt1|a:dk2|a:lt2|a:accent1|a:accent2|a:accent3|a:accent4|a:accent5|a:accent6|a:hlink|a:folHlink", constant_1.theme1File);
    _this.styles["indexedColors"] = _this.readXml.getElementsByTagName("colors/indexedColors/rgbColor", constant_1.stylesFile);
    _this.styles["mruColors"] = _this.readXml.getElementsByTagName("colors/mruColors/color", constant_1.stylesFile);
    _this.imageList = new LuckyImage_1.ImageList(files);

    var numfmts = _this.readXml.getElementsByTagName("numFmt/numFmt", constant_1.stylesFile);

    var numFmtDefaultC = JSON.parse(JSON.stringify(constant_1.numFmtDefault));

    for (var i = 0; i < numfmts.length; i++) {
      var attrList = numfmts[i].attributeList;
      var numfmtid = method_1.getXmlAttibute(attrList, "numFmtId", "49");
      var formatcode = method_1.getXmlAttibute(attrList, "formatCode", "@"); // console.log(numfmtid, formatcode);

      if (!(numfmtid in constant_1.numFmtDefault)) {
        numFmtDefaultC[numfmtid] = constant_1.numFmtDefaultMap[formatcode] || formatcode;
      }
    } // console.log(JSON.stringify(numFmtDefaultC), numfmts);


    _this.styles["numfmts"] = numFmtDefaultC;
    return _this;
  }
  /**
  * @return All sheet name of workbook
  */


  LuckyFile.prototype.getSheetNameList = function () {
    var workbookRelList = this.readXml.getElementsByTagName("Relationships/Relationship", constant_1.workbookRels);

    if (workbookRelList == null) {
      return;
    }

    var regex = new RegExp("worksheets/[^/]*?.xml");
    var sheetNames = {};

    for (var i = 0; i < workbookRelList.length; i++) {
      var rel = workbookRelList[i],
          attrList = rel.attributeList;
      var id = attrList["Id"],
          target = attrList["Target"];

      if (regex.test(target)) {
        if (target.indexOf('/xl') === 0) {
          sheetNames[id] = target.substr(1);
        } else {
          sheetNames[id] = "xl/" + target;
        }
      }
    }

    this.sheetNameList = sheetNames;
  };
  /**
  * @param sheetName WorkSheet'name
  * @return sheet file name and path in zip
  */


  LuckyFile.prototype.getSheetFileBysheetId = function (sheetId) {
    // for(let i=0;i<this.sheetNameList.length;i++){
    //     let sheetFileName = this.sheetNameList[i];
    //     if(sheetFileName.indexOf("sheet"+sheetId)>-1){
    //         return sheetFileName;
    //     }
    // }
    return this.sheetNameList[sheetId];
  };
  /**
  * @return workBook information
  */


  LuckyFile.prototype.getWorkBookInfo = function () {
    var Company = this.readXml.getElementsByTagName("Company", constant_1.appFile);
    var AppVersion = this.readXml.getElementsByTagName("AppVersion", constant_1.appFile);
    var creator = this.readXml.getElementsByTagName("dc:creator", constant_1.coreFile);
    var lastModifiedBy = this.readXml.getElementsByTagName("cp:lastModifiedBy", constant_1.coreFile);
    var created = this.readXml.getElementsByTagName("dcterms:created", constant_1.coreFile);
    var modified = this.readXml.getElementsByTagName("dcterms:modified", constant_1.coreFile);
    this.info = new LuckyBase_1.LuckyFileInfo();
    this.info.name = this.fileName;
    this.info.creator = creator.length > 0 ? creator[0].value : "";
    this.info.lastmodifiedby = lastModifiedBy.length > 0 ? lastModifiedBy[0].value : "";
    this.info.createdTime = created.length > 0 ? created[0].value : "";
    this.info.modifiedTime = modified.length > 0 ? modified[0].value : "";
    this.info.company = Company.length > 0 ? Company[0].value : "";
    this.info.appversion = AppVersion.length > 0 ? AppVersion[0].value : "";
  };
  /**
  * @return All sheet , include whole information
  */


  LuckyFile.prototype.getSheetsFull = function (isInitialCell) {
    if (isInitialCell === void 0) {
      isInitialCell = true;
    }

    var sheets = this.readXml.getElementsByTagName("sheets/sheet", constant_1.workBookFile);
    var sheetList = {};

    for (var key in sheets) {
      var sheet = sheets[key];
      sheetList[sheet.attributeList.name] = sheet.attributeList["sheetId"];
    }

    this.sheets = [];
    var order = 0;

    for (var key in sheets) {
      var sheet = sheets[key];
      var sheetName = sheet.attributeList.name;
      var sheetId = sheet.attributeList["sheetId"];
      var rid = sheet.attributeList["r:id"];
      var sheetFile = this.getSheetFileBysheetId(rid);
      var hide = sheet.attributeList.state === "hidden" ? 1 : 0;
      var drawing = this.readXml.getElementsByTagName("worksheet/drawing", sheetFile),
          drawingFile = void 0,
          drawingRelsFile = void 0;

      if (drawing != null && drawing.length > 0) {
        var attrList = drawing[0].attributeList;
        var rid_1 = method_1.getXmlAttibute(attrList, "r:id", null);

        if (rid_1 != null) {
          drawingFile = this.getDrawingFile(rid_1, sheetFile);
          drawingRelsFile = this.getDrawingRelsFile(drawingFile);
        }
      }

      if (sheetFile != null) {
        var sheet_1 = new LuckySheet_1.LuckySheet(sheetName, sheetId, order, isInitialCell, {
          sheetFile: sheetFile,
          readXml: this.readXml,
          sheetList: sheetList,
          styles: this.styles,
          sharedStrings: this.sharedStrings,
          calcChain: this.calcChain,
          imageList: this.imageList,
          drawingFile: drawingFile,
          drawingRelsFile: drawingRelsFile,
          hide: hide
        });
        this.columnWidthSet = [];
        this.rowHeightSet = [];
        this.imagePositionCaculation(sheet_1);
        this.sheets.push(sheet_1);
        order++;
      }
    }
  };

  LuckyFile.prototype.extendArray = function (index, sets, def, hidden, lens) {
    if (index < sets.length) {
      return;
    }

    var startIndex = sets.length,
        endIndex = index;
    var allGap = 0;

    if (startIndex > 0) {
      allGap = sets[startIndex - 1];
    } // else{
    //     sets.push(0);
    // }


    for (var i = startIndex; i <= endIndex; i++) {
      var gap = def,
          istring = i.toString();

      if (istring in hidden) {
        gap = 0;
      } else if (istring in lens) {
        gap = lens[istring];
      }

      allGap += Math.round(gap + 1);
      sets.push(allGap);
    }
  };

  LuckyFile.prototype.imagePositionCaculation = function (sheet) {
    var images = sheet.images,
        defaultColWidth = sheet.defaultColWidth,
        defaultRowHeight = sheet.defaultRowHeight;
    var colhidden = {};

    if (sheet.config.colhidden) {
      colhidden = sheet.config.colhidden;
    }

    var columnlen = {};

    if (sheet.config.columnlen) {
      columnlen = sheet.config.columnlen;
    }

    var rowhidden = {};

    if (sheet.config.rowhidden) {
      rowhidden = sheet.config.rowhidden;
    }

    var rowlen = {};

    if (sheet.config.rowlen) {
      rowlen = sheet.config.rowlen;
    }

    for (var key in images) {
      var imageObject = images[key]; //Image, luckyImage

      var fromCol = imageObject.fromCol;
      var fromColOff = imageObject.fromColOff;
      var fromRow = imageObject.fromRow;
      var fromRowOff = imageObject.fromRowOff;
      var toCol = imageObject.toCol;
      var toColOff = imageObject.toColOff;
      var toRow = imageObject.toRow;
      var toRowOff = imageObject.toRowOff;
      var x_n = 0,
          y_n = 0;
      var cx_n = 0,
          cy_n = 0;

      if (fromCol >= this.columnWidthSet.length) {
        this.extendArray(fromCol, this.columnWidthSet, defaultColWidth, colhidden, columnlen);
      }

      if (fromCol == 0) {
        x_n = 0;
      } else {
        x_n = this.columnWidthSet[fromCol - 1];
      }

      x_n = x_n + fromColOff;

      if (fromRow >= this.rowHeightSet.length) {
        this.extendArray(fromRow, this.rowHeightSet, defaultRowHeight, rowhidden, rowlen);
      }

      if (fromRow == 0) {
        y_n = 0;
      } else {
        y_n = this.rowHeightSet[fromRow - 1];
      }

      y_n = y_n + fromRowOff;

      if (toCol >= this.columnWidthSet.length) {
        this.extendArray(toCol, this.columnWidthSet, defaultColWidth, colhidden, columnlen);
      }

      if (toCol == 0) {
        cx_n = 0;
      } else {
        cx_n = this.columnWidthSet[toCol - 1];
      }

      cx_n = cx_n + toColOff - x_n;

      if (toRow >= this.rowHeightSet.length) {
        this.extendArray(toRow, this.rowHeightSet, defaultRowHeight, rowhidden, rowlen);
      }

      if (toRow == 0) {
        cy_n = 0;
      } else {
        cy_n = this.rowHeightSet[toRow - 1];
      }

      cy_n = cy_n + toRowOff - y_n;
      console.log(defaultColWidth, colhidden, columnlen);
      console.log(fromCol, this.columnWidthSet[fromCol], fromColOff);
      console.log(toCol, this.columnWidthSet[toCol], toColOff, JSON.stringify(this.columnWidthSet));
      imageObject.originWidth = cx_n;
      imageObject.originHeight = cy_n;
      imageObject.crop.height = cy_n;
      imageObject.crop.width = cx_n;
      imageObject["default"].height = cy_n;
      imageObject["default"].left = x_n;
      imageObject["default"].top = y_n;
      imageObject["default"].width = cx_n;
    }

    console.log(this.columnWidthSet, this.rowHeightSet);
  };
  /**
  * @return drawing file string
  */


  LuckyFile.prototype.getDrawingFile = function (rid, sheetFile) {
    var sheetRelsPath = "xl/worksheets/_rels/";
    var sheetFileArr = sheetFile.split("/");
    var sheetRelsName = sheetFileArr[sheetFileArr.length - 1];
    var sheetRelsFile = sheetRelsPath + sheetRelsName + ".rels";
    var drawing = this.readXml.getElementsByTagName("Relationships/Relationship", sheetRelsFile);

    if (drawing.length > 0) {
      for (var i = 0; i < drawing.length; i++) {
        var relationship = drawing[i];
        var attrList = relationship.attributeList;
        var relationshipId = method_1.getXmlAttibute(attrList, "Id", null);

        if (relationshipId == rid) {
          var target = method_1.getXmlAttibute(attrList, "Target", null);

          if (target != null) {
            return target.replace(/\.\.\//g, "");
          }
        }
      }
    }

    return null;
  };

  LuckyFile.prototype.getDrawingRelsFile = function (drawingFile) {
    var drawingRelsPath = "xl/drawings/_rels/";
    var drawingFileArr = drawingFile.split("/");
    var drawingRelsName = drawingFileArr[drawingFileArr.length - 1];
    var drawingRelsFile = drawingRelsPath + drawingRelsName + ".rels";
    return drawingRelsFile;
  };
  /**
  * @return All sheet base information widthout cell and config
  */


  LuckyFile.prototype.getSheetsWithoutCell = function () {
    this.getSheetsFull(false);
  };
  /**
  * @return LuckySheet file json
  */


  LuckyFile.prototype.Parse = function () {
    // let xml = this.readXml;
    // for(let key in this.sheetNameList){
    //     let sheetName=this.sheetNameList[key];
    //     let sheetColumns = xml.getElementsByTagName("row/c/f", sheetName);
    //     console.log(sheetColumns);
    // }
    // return "";
    this.getWorkBookInfo();
    this.getSheetsFull(); // for(let i=0;i<this.sheets.length;i++){
    //     let sheet = this.sheets[i];
    //     let _borderInfo = sheet.config._borderInfo;
    //     if(_borderInfo==null){
    //         continue;
    //     }
    //     let _borderInfoKeys = Object.keys(_borderInfo);
    //     _borderInfoKeys.sort();
    //     for(let a=0;a<_borderInfoKeys.length;a++){
    //         let key = parseInt(_borderInfoKeys[a]);
    //         let b = _borderInfo[key];
    //         if(b.cells.length==0){
    //             continue;
    //         }
    //         if(sheet.config.borderInfo==null){
    //             sheet.config.borderInfo = [];
    //         }
    //         sheet.config.borderInfo.push(b);
    //     }
    // }

    return this.toJsonString(this);
  };

  LuckyFile.prototype.toJsonString = function (file) {
    var LuckyOutPutFile = new LuckyBase_1.LuckyFileBase();
    LuckyOutPutFile.info = file.info;
    LuckyOutPutFile.sheets = [];
    file.sheets.forEach(function (sheet) {
      var sheetout = new LuckyBase_1.LuckySheetBase(); //let attrName = ["name","color","config","index","status","order","row","column","luckysheet_select_save","scrollLeft","scrollTop","zoomRatio","showGridLines","defaultColWidth","defaultRowHeight","celldata","chart","isPivotTable","pivotTable","luckysheet_conditionformat_save","freezen","calcChain"];

      if (sheet.name != null) {
        sheetout.name = sheet.name;
      }

      if (sheet.color != null) {
        sheetout.color = sheet.color;
      }

      if (sheet.config != null) {
        sheetout.config = sheet.config; // if(sheetout.config._borderInfo!=null){
        //     delete sheetout.config._borderInfo;
        // }
      }

      if (sheet.index != null) {
        sheetout.index = sheet.index;
      }

      if (sheet.status != null) {
        sheetout.status = sheet.status;
      }

      if (sheet.order != null) {
        sheetout.order = sheet.order;
      }

      if (sheet.row != null) {
        sheetout.row = sheet.row;
      }

      if (sheet.column != null) {
        sheetout.column = sheet.column;
      }

      if (sheet.luckysheet_select_save != null) {
        sheetout.luckysheet_select_save = sheet.luckysheet_select_save;
      }

      if (sheet.scrollLeft != null) {
        sheetout.scrollLeft = sheet.scrollLeft;
      }

      if (sheet.scrollTop != null) {
        sheetout.scrollTop = sheet.scrollTop;
      }

      if (sheet.zoomRatio != null) {
        sheetout.zoomRatio = sheet.zoomRatio;
      }

      if (sheet.showGridLines != null) {
        sheetout.showGridLines = sheet.showGridLines;
      }

      if (sheet.defaultColWidth != null) {
        sheetout.defaultColWidth = sheet.defaultColWidth;
      }

      if (sheet.defaultRowHeight != null) {
        sheetout.defaultRowHeight = sheet.defaultRowHeight;
      }

      if (sheet.celldata != null) {
        // sheetout.celldata = sheet.celldata;
        sheetout.celldata = [];
        sheet.celldata.forEach(function (cell) {
          var cellout = new LuckyBase_1.LuckySheetCelldataBase();
          cellout.r = cell.r;
          cellout.c = cell.c;
          cellout.v = cell.v;
          sheetout.celldata.push(cellout);
        });
      }

      if (sheet.chart != null) {
        sheetout.chart = sheet.chart;
      }

      if (sheet.isPivotTable != null) {
        sheetout.isPivotTable = sheet.isPivotTable;
      }

      if (sheet.pivotTable != null) {
        sheetout.pivotTable = sheet.pivotTable;
      }

      if (sheet.luckysheet_conditionformat_save != null) {
        sheetout.luckysheet_conditionformat_save = sheet.luckysheet_conditionformat_save;
      }

      if (sheet.freezen != null) {
        sheetout.freezen = sheet.freezen;
      }

      if (sheet.calcChain != null) {
        sheetout.calcChain = sheet.calcChain;
      }

      if (sheet.images != null) {
        sheetout.images = sheet.images;
      }

      if (sheet.dataVerification != null) {
        sheetout.dataVerification = sheet.dataVerification;
      }

      if (sheet.hyperlink != null) {
        sheetout.hyperlink = sheet.hyperlink;
      }

      if (sheet.hide != null) {
        sheetout.hide = sheet.hide;
      }

      LuckyOutPutFile.sheets.push(sheetout);
    });
    return JSON.stringify(LuckyOutPutFile);
  };

  return LuckyFile;
}(LuckyBase_1.LuckyFileBase);

exports.LuckyFile = LuckyFile;

},{"../common/constant":16,"../common/method":18,"./LuckyBase":10,"./LuckyImage":13,"./LuckySheet":14,"./ReadXml":15}],13:[function(require,module,exports){
"use strict";

var __extends = void 0 && (void 0).__extends || function () {
  var _extendStatics = function extendStatics(d, b) {
    _extendStatics = Object.setPrototypeOf || {
      __proto__: []
    } instanceof Array && function (d, b) {
      d.__proto__ = b;
    } || function (d, b) {
      for (var p in b) {
        if (b.hasOwnProperty(p)) d[p] = b[p];
      }
    };

    return _extendStatics(d, b);
  };

  return function (d, b) {
    _extendStatics(d, b);

    function __() {
      this.constructor = d;
    }

    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
  };
}();

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.ImageList = void 0;

var LuckyBase_1 = require("./LuckyBase");

var emf_1 = require("../common/emf");

var ImageList =
/** @class */
function () {
  function ImageList(files) {
    if (files == null) {
      return;
    }

    this.images = {};

    for (var fileKey in files) {
      // let reg = new RegExp("xl/media/image1.png", "g");
      if (fileKey.indexOf("xl/media/") > -1) {
        var fileNameArr = fileKey.split(".");
        var suffix = fileNameArr[fileNameArr.length - 1].toLowerCase();

        if (suffix in {
          "png": 1,
          "jpeg": 1,
          "jpg": 1,
          "gif": 1,
          "bmp": 1,
          "tif": 1,
          "webp": 1,
          "emf": 1
        }) {
          if (suffix == "emf") {
            var pNum = 0; // number of the page, that you want to render

            var scale = 1; // the scale of the document

            var wrt = new emf_1.ToContext2D(pNum, scale);
            var inp, out, stt;
            emf_1.FromEMF.K = [];
            inp = emf_1.FromEMF.C;
            out = emf_1.FromEMF.K;
            stt = 4;

            for (var p in inp) {
              out[inp[p]] = p.slice(stt);
            }

            emf_1.FromEMF.Parse(files[fileKey], wrt);
            this.images[fileKey] = wrt.canvas.toDataURL("image/png");
          } else {
            this.images[fileKey] = files[fileKey];
          }
        }
      }
    }
  }

  ImageList.prototype.getImageByName = function (pathName) {
    if (pathName in this.images) {
      var base64 = this.images[pathName];
      return new Image(pathName, base64);
    }

    return null;
  };

  return ImageList;
}();

exports.ImageList = ImageList;

var Image =
/** @class */
function (_super) {
  __extends(Image, _super);

  function Image(pathName, base64) {
    var _this = _super.call(this) || this;

    _this.src = base64;
    return _this;
  }

  Image.prototype.setDefault = function () {};

  return Image;
}(LuckyBase_1.LuckyImageBase);

},{"../common/emf":17,"./LuckyBase":10}],14:[function(require,module,exports){
"use strict";

var __extends = void 0 && (void 0).__extends || function () {
  var _extendStatics = function extendStatics(d, b) {
    _extendStatics = Object.setPrototypeOf || {
      __proto__: []
    } instanceof Array && function (d, b) {
      d.__proto__ = b;
    } || function (d, b) {
      for (var p in b) {
        if (b.hasOwnProperty(p)) d[p] = b[p];
      }
    };

    return _extendStatics(d, b);
  };

  return function (d, b) {
    _extendStatics(d, b);

    function __() {
      this.constructor = d;
    }

    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
  };
}();

var __importDefault = void 0 && (void 0).__importDefault || function (mod) {
  return mod && mod.__esModule ? mod : {
    "default": mod
  };
};

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.LuckySheet = void 0;

var LuckyCell_1 = require("./LuckyCell");

var method_1 = require("../common/method");

var constant_1 = require("../common/constant");

var ReadXml_1 = require("./ReadXml");

var LuckyBase_1 = require("./LuckyBase");

var dayjs_1 = __importDefault(require("dayjs"));

var LuckySheet =
/** @class */
function (_super) {
  __extends(LuckySheet, _super);

  function LuckySheet(sheetName, sheetId, sheetOrder, isInitialCell, allFileOption) {
    if (isInitialCell === void 0) {
      isInitialCell = false;
    }

    var _this = //Private
    _super.call(this) || this;

    _this.isInitialCell = isInitialCell;
    _this.readXml = allFileOption.readXml;
    _this.sheetFile = allFileOption.sheetFile;
    _this.styles = allFileOption.styles;
    _this.sharedStrings = allFileOption.sharedStrings;
    _this.calcChainEles = allFileOption.calcChain;
    _this.sheetList = allFileOption.sheetList;
    _this.imageList = allFileOption.imageList;
    _this.hide = allFileOption.hide; //Output

    _this.name = sheetName;
    _this.index = sheetId;
    _this.order = sheetOrder.toString();
    _this.config = new LuckyBase_1.LuckyConfig();
    _this.celldata = [];
    _this.mergeCells = _this.readXml.getElementsByTagName("mergeCells/mergeCell", _this.sheetFile);
    var clrScheme = _this.styles["clrScheme"];

    var sheetView = _this.readXml.getElementsByTagName("sheetViews/sheetView", _this.sheetFile);

    var showGridLines = "1",
        tabSelected = "0",
        zoomScale = "100",
        activeCell = "A1";

    if (sheetView.length > 0) {
      var attrList = sheetView[0].attributeList;
      showGridLines = method_1.getXmlAttibute(attrList, "showGridLines", "1");
      tabSelected = method_1.getXmlAttibute(attrList, "tabSelected", "0");
      zoomScale = method_1.getXmlAttibute(attrList, "zoomScale", "100"); // let colorId = getXmlAttibute(attrList, "colorId", "0");

      var selections = sheetView[0].getInnerElements("selection");

      if (selections != null && selections.length > 0) {
        activeCell = method_1.getXmlAttibute(selections[0].attributeList, "activeCell", "A1");
        var range = method_1.getcellrange(activeCell, _this.sheetList, sheetId);
        _this.luckysheet_select_save = [];

        _this.luckysheet_select_save.push(range);
      }
    }

    _this.showGridLines = showGridLines;
    _this.status = tabSelected;
    _this.zoomRatio = parseInt(zoomScale) / 100;

    var tabColors = _this.readXml.getElementsByTagName("sheetPr/tabColor", _this.sheetFile);

    if (tabColors != null && tabColors.length > 0) {
      var tabColor = tabColors[0],
          attrList = tabColor.attributeList; // if(attrList.rgb!=null){

      var tc = ReadXml_1.getColor(tabColor, _this.styles, "b");
      _this.color = tc; // }
    }

    var sheetFormatPr = _this.readXml.getElementsByTagName("sheetFormatPr", _this.sheetFile);

    var defaultColWidth, defaultRowHeight;

    if (sheetFormatPr.length > 0) {
      var attrList = sheetFormatPr[0].attributeList;
      defaultColWidth = method_1.getXmlAttibute(attrList, "defaultColWidth", "9.21");
      defaultRowHeight = method_1.getXmlAttibute(attrList, "defaultRowHeight", "19");
    }

    _this.defaultColWidth = method_1.getColumnWidthPixel(parseFloat(defaultColWidth));
    _this.defaultRowHeight = method_1.getRowHeightPixel(parseFloat(defaultRowHeight));

    _this.generateConfigColumnLenAndHidden();

    var cellOtherInfo = _this.generateConfigRowLenAndHiddenAddCell();

    if (_this.formulaRefList != null) {
      for (var key in _this.formulaRefList) {
        var funclist = _this.formulaRefList[key];
        var mainFunc = funclist["mainRef"],
            mainCellValue = mainFunc.cellValue;
        var formulaTxt = mainFunc.fv;
        var mainR = mainCellValue.r,
            mainC = mainCellValue.c; // let refRange = getcellrange(ref);

        for (var name_1 in funclist) {
          if (name_1 == "mainRef") {
            continue;
          }

          var funcValue = funclist[name_1],
              cellValue = funcValue.cellValue;

          if (cellValue == null) {
            continue;
          }

          var r = cellValue.r,
              c = cellValue.c;
          var func = formulaTxt;
          var offsetRow = r - mainR,
              offsetCol = c - mainC;

          if (offsetRow > 0) {
            func = "=" + method_1.fromulaRef.functionCopy(func, "down", offsetRow);
          } else if (offsetRow < 0) {
            func = "=" + method_1.fromulaRef.functionCopy(func, "up", Math.abs(offsetRow));
          }

          if (offsetCol > 0) {
            func = "=" + method_1.fromulaRef.functionCopy(func, "right", offsetCol);
          } else if (offsetCol < 0) {
            func = "=" + method_1.fromulaRef.functionCopy(func, "left", Math.abs(offsetCol));
          } // console.log(offsetRow, offsetCol, func);


          cellValue.v.f = func;
        }
      }
    }

    if (_this.calcChain == null) {
      _this.calcChain = [];
    }

    var formulaListExist = {};

    for (var c = 0; c < _this.calcChainEles.length; c++) {
      var calcChainEle = _this.calcChainEles[c],
          attrList = calcChainEle.attributeList;

      if (attrList.i != sheetId) {
        continue;
      }

      var r = attrList.r,
          i = attrList.i,
          l = attrList.l,
          s = attrList.s,
          a = attrList.a,
          t = attrList.t;
      var range = method_1.getcellrange(r);
      var chain = new LuckyBase_1.LuckysheetCalcChain();
      chain.r = range.row[0];
      chain.c = range.column[0];
      chain.index = _this.index;

      _this.calcChain.push(chain);

      formulaListExist["r" + r + "c" + c] = null;
    } //There may be formulas that do not appear in calcChain


    for (var key in cellOtherInfo.formulaList) {
      if (!(key in formulaListExist)) {
        var formulaListItem = cellOtherInfo.formulaList[key];
        var chain = new LuckyBase_1.LuckysheetCalcChain();
        chain.r = formulaListItem.r;
        chain.c = formulaListItem.c;
        chain.index = _this.index;

        _this.calcChain.push(chain);
      }
    } // dataVerification config


    _this.dataVerification = _this.generateConfigDataValidations(); // hyperlink config

    _this.hyperlink = _this.generateConfigHyperlinks(); // sheet hide

    _this.hide = _this.hide;

    if (_this.mergeCells != null) {
      for (var i = 0; i < _this.mergeCells.length; i++) {
        var merge = _this.mergeCells[i],
            attrList = merge.attributeList;
        var ref = attrList.ref;

        if (ref == null) {
          continue;
        }

        var range = method_1.getcellrange(ref, _this.sheetList, sheetId);
        var mergeValue = new LuckyBase_1.LuckySheetConfigMerge();
        mergeValue.r = range.row[0];
        mergeValue.c = range.column[0];
        mergeValue.rs = range.row[1] - range.row[0] + 1;
        mergeValue.cs = range.column[1] - range.column[0] + 1;

        if (_this.config.merge == null) {
          _this.config.merge = {};
        }

        _this.config.merge[range.row[0] + "_" + range.column[0]] = mergeValue;
      }
    }

    var drawingFile = allFileOption.drawingFile,
        drawingRelsFile = allFileOption.drawingRelsFile;

    if (drawingFile != null && drawingRelsFile != null) {
      var twoCellAnchors = _this.readXml.getElementsByTagName("xdr:twoCellAnchor", drawingFile);

      if (twoCellAnchors != null && twoCellAnchors.length > 0) {
        for (var i = 0; i < twoCellAnchors.length; i++) {
          var twoCellAnchor = twoCellAnchors[i];
          var editAs = method_1.getXmlAttibute(twoCellAnchor.attributeList, "editAs", "twoCell");
          var xdrFroms = twoCellAnchor.getInnerElements("xdr:from"),
              xdrTos = twoCellAnchor.getInnerElements("xdr:to");
          var xdr_blipfills = twoCellAnchor.getInnerElements("a:blip");

          if (xdrFroms != null && xdr_blipfills != null && xdrFroms.length > 0 && xdr_blipfills.length > 0) {
            var xdrFrom = xdrFroms[0],
                xdrTo = xdrTos[0],
                xdr_blipfill = xdr_blipfills[0];
            var rembed = method_1.getXmlAttibute(xdr_blipfill.attributeList, "r:embed", null);

            var imageObject = _this.getBase64ByRid(rembed, drawingRelsFile); // let aoff = xdr_xfrm.getInnerElements("a:off"), aext = xdr_xfrm.getInnerElements("a:ext");
            // if(aoff!=null && aext!=null && aoff.length>0 && aext.length>0){
            //     let aoffAttribute = aoff[0].attributeList, aextAttribute = aext[0].attributeList;
            //     let x = getXmlAttibute(aoffAttribute, "x", null);
            //     let y = getXmlAttibute(aoffAttribute, "y", null);
            //     let cx = getXmlAttibute(aextAttribute, "cx", null);
            //     let cy = getXmlAttibute(aextAttribute, "cy", null);
            //     if(x!=null && y!=null && cx!=null && cy!=null && imageObject !=null){
            // let x_n = getPxByEMUs(parseInt(x), "c"),y_n = getPxByEMUs(parseInt(y));
            // let cx_n = getPxByEMUs(parseInt(cx), "c"),cy_n = getPxByEMUs(parseInt(cy));


            var x_n = 0,
                y_n = 0;
            var cx_n = 0,
                cy_n = 0;
            imageObject.fromCol = _this.getXdrValue(xdrFrom.getInnerElements("xdr:col"));
            imageObject.fromColOff = method_1.getPxByEMUs(_this.getXdrValue(xdrFrom.getInnerElements("xdr:colOff")));
            imageObject.fromRow = _this.getXdrValue(xdrFrom.getInnerElements("xdr:row"));
            imageObject.fromRowOff = method_1.getPxByEMUs(_this.getXdrValue(xdrFrom.getInnerElements("xdr:rowOff")));
            imageObject.toCol = _this.getXdrValue(xdrTo.getInnerElements("xdr:col"));
            imageObject.toColOff = method_1.getPxByEMUs(_this.getXdrValue(xdrTo.getInnerElements("xdr:colOff")));
            imageObject.toRow = _this.getXdrValue(xdrTo.getInnerElements("xdr:row"));
            imageObject.toRowOff = method_1.getPxByEMUs(_this.getXdrValue(xdrTo.getInnerElements("xdr:rowOff")));
            imageObject.originWidth = cx_n;
            imageObject.originHeight = cy_n;

            if (editAs == "absolute") {
              imageObject.type = "3";
            } else if (editAs == "oneCell") {
              imageObject.type = "2";
            } else {
              imageObject.type = "1";
            }

            imageObject.isFixedPos = false;
            imageObject.fixedLeft = 0;
            imageObject.fixedTop = 0;
            var imageBorder = {
              color: "#000",
              radius: 0,
              style: "solid",
              width: 0
            };
            imageObject.border = imageBorder;
            var imageCrop = {
              height: cy_n,
              offsetLeft: 0,
              offsetTop: 0,
              width: cx_n
            };
            imageObject.crop = imageCrop;
            var imageDefault = {
              height: cy_n,
              left: x_n,
              top: y_n,
              width: cx_n
            };
            imageObject["default"] = imageDefault;

            if (_this.images == null) {
              _this.images = {};
            }

            _this.images[method_1.generateRandomIndex("image")] = imageObject; //     }
            // }
          }
        }
      }
    }

    return _this;
  }

  LuckySheet.prototype.getXdrValue = function (ele) {
    if (ele == null || ele.length == 0) {
      return null;
    }

    return parseInt(ele[0].value);
  };

  LuckySheet.prototype.getBase64ByRid = function (rid, drawingRelsFile) {
    var Relationships = this.readXml.getElementsByTagName("Relationships/Relationship", drawingRelsFile);

    if (Relationships != null && Relationships.length > 0) {
      for (var i = 0; i < Relationships.length; i++) {
        var Relationship = Relationships[i];
        var attrList = Relationship.attributeList;
        var Id = method_1.getXmlAttibute(attrList, "Id", null);
        var src = method_1.getXmlAttibute(attrList, "Target", null);

        if (Id == rid) {
          src = src.replace(/\.\.\//g, "");
          src = "xl/" + src;
          var imgage = this.imageList.getImageByName(src);
          return imgage;
        }
      }
    }

    return null;
  };
  /**
  * @desc This will convert cols/col to luckysheet config of column'width
  */


  LuckySheet.prototype.generateConfigColumnLenAndHidden = function () {
    var cols = this.readXml.getElementsByTagName("cols/col", this.sheetFile);

    for (var i = 0; i < cols.length; i++) {
      var col = cols[i],
          attrList = col.attributeList;
      var min = method_1.getXmlAttibute(attrList, "min", null);
      var max = method_1.getXmlAttibute(attrList, "max", null);
      var width = method_1.getXmlAttibute(attrList, "width", null);
      var hidden = method_1.getXmlAttibute(attrList, "hidden", null);
      var customWidth = method_1.getXmlAttibute(attrList, "customWidth", null);

      if (min == null || max == null) {
        continue;
      }

      var minNum = parseInt(min) - 1,
          maxNum = parseInt(max) - 1,
          widthNum = parseFloat(width);

      for (var m = minNum; m <= maxNum; m++) {
        if (width != null) {
          if (this.config.columnlen == null) {
            this.config.columnlen = {};
          }

          this.config.columnlen[m] = method_1.getColumnWidthPixel(widthNum);
        }

        if (hidden == "1") {
          if (this.config.colhidden == null) {
            this.config.colhidden = {};
          }

          this.config.colhidden[m] = 0;

          if (this.config.columnlen) {
            delete this.config.columnlen[m];
          }
        }

        if (customWidth != null) {
          if (this.config.customWidth == null) {
            this.config.customWidth = {};
          }

          this.config.customWidth[m] = 1;
        }
      }
    }
  };
  /**
  * @desc This will convert cols/col to luckysheet config of column'width
  */


  LuckySheet.prototype.generateConfigRowLenAndHiddenAddCell = function () {
    var rows = this.readXml.getElementsByTagName("sheetData/row", this.sheetFile);
    var cellOtherInfo = {};
    var formulaList = {};
    cellOtherInfo.formulaList = formulaList;

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i],
          attrList = row.attributeList;
      var rowNo = method_1.getXmlAttibute(attrList, "r", null);
      var height = method_1.getXmlAttibute(attrList, "ht", null);
      var hidden = method_1.getXmlAttibute(attrList, "hidden", null);
      var customHeight = method_1.getXmlAttibute(attrList, "customHeight", null);

      if (rowNo == null) {
        continue;
      }

      var rowNoNum = parseInt(rowNo) - 1;

      if (height != null) {
        var heightNum = parseFloat(height);

        if (this.config.rowlen == null) {
          this.config.rowlen = {};
        }

        this.config.rowlen[rowNoNum] = method_1.getRowHeightPixel(heightNum);
      }

      if (hidden == "1") {
        if (this.config.rowhidden == null) {
          this.config.rowhidden = {};
        }

        this.config.rowhidden[rowNoNum] = 0;

        if (this.config.rowlen) {
          delete this.config.rowlen[rowNoNum];
        }
      }

      if (customHeight != null) {
        if (this.config.customHeight == null) {
          this.config.customHeight = {};
        }

        this.config.customHeight[rowNoNum] = 1;
      }

      if (this.isInitialCell) {
        var cells = row.getInnerElements("c");

        for (var key in cells) {
          var cell = cells[key];
          var cellValue = new LuckyCell_1.LuckySheetCelldata(cell, this.styles, this.sharedStrings, this.mergeCells, this.sheetFile, this.readXml);

          if (cellValue._borderObject != null) {
            if (this.config.borderInfo == null) {
              this.config.borderInfo = [];
            }

            this.config.borderInfo.push(cellValue._borderObject);
            delete cellValue._borderObject;
          } // let borderId = cellValue._borderId;
          // if(borderId!=null){
          //     let borders = this.styles["borders"] as Element[];
          //     if(this.config._borderInfo==null){
          //         this.config._borderInfo = {};
          //     }
          //     if( borderId in this.config._borderInfo){
          //         this.config._borderInfo[borderId].cells.push(cellValue.r + "_" + cellValue.c);
          //     }
          //     else{
          //         let border = borders[borderId];
          //         let borderObject = new LuckySheetborderInfoCellForImp();
          //         borderObject.rangeType = "cellGroup";
          //         borderObject.cells = [];
          //         let borderCellValue = new LuckySheetborderInfoCellValue();
          //         let lefts = border.getInnerElements("left");
          //         let rights = border.getInnerElements("right");
          //         let tops = border.getInnerElements("top");
          //         let bottoms = border.getInnerElements("bottom");
          //         let diagonals = border.getInnerElements("diagonal");
          //         let left = this.getBorderInfo(lefts);
          //         let right = this.getBorderInfo(rights);
          //         let top = this.getBorderInfo(tops);
          //         let bottom = this.getBorderInfo(bottoms);
          //         let diagonal = this.getBorderInfo(diagonals);
          //         let isAdd = false;
          //         if(left!=null && left.color!=null){
          //             borderCellValue.l = left;
          //             isAdd = true;
          //         }
          //         if(right!=null && right.color!=null){
          //             borderCellValue.r = right;
          //             isAdd = true;
          //         }
          //         if(top!=null && top.color!=null){
          //             borderCellValue.t = top;
          //             isAdd = true;
          //         }
          //         if(bottom!=null && bottom.color!=null){
          //             borderCellValue.b = bottom;
          //             isAdd = true;
          //         }
          //         if(isAdd){
          //             borderObject.value = borderCellValue;
          //             this.config._borderInfo[borderId] = borderObject;
          //         }
          //     }
          // }


          if (cellValue._formulaType == "shared") {
            if (this.formulaRefList == null) {
              this.formulaRefList = {};
            }

            if (this.formulaRefList[cellValue._formulaSi] == null) {
              this.formulaRefList[cellValue._formulaSi] = {};
            }

            var fv = void 0;

            if (cellValue.v != null) {
              fv = cellValue.v.f;
            }

            var refValue = {
              t: cellValue._formulaType,
              ref: cellValue._fomulaRef,
              si: cellValue._formulaSi,
              fv: fv,
              cellValue: cellValue
            };

            if (cellValue._fomulaRef != null) {
              this.formulaRefList[cellValue._formulaSi]["mainRef"] = refValue;
            } else {
              this.formulaRefList[cellValue._formulaSi][cellValue.r + "_" + cellValue.c] = refValue;
            } // console.log(refValue, this.formulaRefList);

          } //There may be formulas that do not appear in calcChain


          if (cellValue.v != null && cellValue.v.f != null) {
            var formulaCell = {
              r: cellValue.r,
              c: cellValue.c
            };
            cellOtherInfo.formulaList["r" + cellValue.r + "c" + cellValue.c] = formulaCell;
          }

          this.celldata.push(cellValue);
        }
      }
    }

    return cellOtherInfo;
  };
  /**
   * luckysheet config of dataValidations
   *
   * @returns {IluckysheetDataVerification} - dataValidations config
   */


  LuckySheet.prototype.generateConfigDataValidations = function () {
    var rows = this.readXml.getElementsByTagName("dataValidations/dataValidation", this.sheetFile);
    var extLst = this.readXml.getElementsByTagName("extLst/ext/x14:dataValidations/x14:dataValidation", this.sheetFile) || [];
    rows = rows.concat(extLst);
    var dataVerification = {};

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var attrList = row.attributeList;
      var formulaValue = row.value;
      var type = method_1.getXmlAttibute(attrList, "type", null);
      var operator = "",
          sqref = "",
          sqrefIndexArr = [],
          valueArr = [];

      var _prohibitInput = method_1.getXmlAttibute(attrList, "allowBlank", null) !== "1" ? false : true; // x14 processing


      var formulaReg = new RegExp(/<x14:formula1>|<xm:sqref>/g);

      if (formulaReg.test(formulaValue) && (extLst === null || extLst === void 0 ? void 0 : extLst.length) >= 0) {
        operator = method_1.getXmlAttibute(attrList, "operator", null);
        var peelOffData = method_1.getPeelOffX14(formulaValue);
        sqref = peelOffData === null || peelOffData === void 0 ? void 0 : peelOffData.sqref;
        sqrefIndexArr = method_1.getMultiSequenceToNum(sqref);
        valueArr = method_1.getMultiFormulaValue(peelOffData === null || peelOffData === void 0 ? void 0 : peelOffData.formula);
      } else {
        operator = method_1.getXmlAttibute(attrList, "operator", null);
        sqref = method_1.getXmlAttibute(attrList, "sqref", null);
        sqrefIndexArr = method_1.getMultiSequenceToNum(sqref);
        valueArr = method_1.getMultiFormulaValue(formulaValue);
      }

      var _type = constant_1.DATA_VERIFICATION_MAP[type];
      var _type2 = null;

      var _value1 = (valueArr === null || valueArr === void 0 ? void 0 : valueArr.length) >= 1 ? valueArr[0] : "";

      var _value2 = (valueArr === null || valueArr === void 0 ? void 0 : valueArr.length) === 2 ? valueArr[1] : "";

      var _hint = method_1.getXmlAttibute(attrList, "prompt", null);

      var _hintShow = _hint ? true : false;

      var matchType = constant_1.COMMON_TYPE2.includes(_type) ? "common" : _type;
      _type2 = operator ? constant_1.DATA_VERIFICATION_TYPE2_MAP[matchType][operator] : "bw"; // mobile phone number processing

      if (_type === "text_content" && ((_value1 === null || _value1 === void 0 ? void 0 : _value1.includes("LEN")) || (_value1 === null || _value1 === void 0 ? void 0 : _value1.includes("len"))) && (_value1 === null || _value1 === void 0 ? void 0 : _value1.includes("=11"))) {
        _type = "validity";
        _type2 = "phone";
      } // date processing


      if (_type === "date") {
        var D1900 = new Date(1899, 11, 30, 0, 0, 0);
        _value1 = dayjs_1["default"](D1900).clone().add(Number(_value1), "day").format("YYYY-MM-DD");
        _value2 = dayjs_1["default"](D1900).clone().add(Number(_value2), "day").format("YYYY-MM-DD");
      } // checkbox and dropdown processing


      if (_type === "checkbox" || _type === "dropdown") {
        _type2 = null;
      } // dynamically add dataVerifications


      for (var _i = 0, sqrefIndexArr_1 = sqrefIndexArr; _i < sqrefIndexArr_1.length; _i++) {
        var ref = sqrefIndexArr_1[_i];
        dataVerification[ref] = {
          type: _type,
          type2: _type2,
          value1: _value1,
          value2: _value2,
          checked: false,
          remote: false,
          prohibitInput: _prohibitInput,
          hintShow: _hintShow,
          hintText: _hint
        };
      }
    }

    return dataVerification;
  };
  /**
   * luckysheet config of hyperlink
   *
   * @returns {IluckysheetHyperlink} - hyperlink config
   */


  LuckySheet.prototype.generateConfigHyperlinks = function () {
    var _a;

    var rows = this.readXml.getElementsByTagName("hyperlinks/hyperlink", this.sheetFile);
    var hyperlink = {};

    var _loop_1 = function _loop_1(i) {
      var row = rows[i];
      var attrList = row.attributeList;

      var ref = method_1.getXmlAttibute(attrList, "ref", null),
          refArr = method_1.getMultiSequenceToNum(ref),
          _display = method_1.getXmlAttibute(attrList, "display", null),
          _address = method_1.getXmlAttibute(attrList, "location", null),
          _tooltip = method_1.getXmlAttibute(attrList, "tooltip", null);

      var _type = _address ? "internal" : "external"; // external hyperlink


      if (!_address) {
        var rid_1 = attrList["r:id"];
        var sheetFile = this_1.sheetFile;
        var relationshipList = this_1.readXml.getElementsByTagName("Relationships/Relationship", "xl/worksheets/_rels/" + sheetFile.replace(constant_1.worksheetFilePath, "") + ".rels");
        var findRid = relationshipList === null || relationshipList === void 0 ? void 0 : relationshipList.find(function (e) {
          return e.attributeList["Id"] === rid_1;
        });

        if (findRid) {
          _address = findRid.attributeList["Target"];
          _type = (_a = findRid.attributeList["TargetMode"]) === null || _a === void 0 ? void 0 : _a.toLocaleLowerCase();
        }
      } // match R1C1


      var addressReg = new RegExp(/^.*!R([\d$])+C([\d$])*$/g);

      if (addressReg.test(_address)) {
        _address = method_1.getTransR1C1ToSequence(_address);
      } // dynamically add hyperlinks


      for (var _i = 0, refArr_1 = refArr; _i < refArr_1.length; _i++) {
        var ref_1 = refArr_1[_i];
        hyperlink[ref_1] = {
          linkAddress: _address,
          linkTooltip: _tooltip || "",
          linkType: _type,
          display: _display || ""
        };
      }
    };

    var this_1 = this;

    for (var i = 0; i < rows.length; i++) {
      _loop_1(i);
    }

    return hyperlink;
  };

  return LuckySheet;
}(LuckyBase_1.LuckySheetBase);

exports.LuckySheet = LuckySheet;

},{"../common/constant":16,"../common/method":18,"./LuckyBase":10,"./LuckyCell":11,"./ReadXml":15,"dayjs":3}],15:[function(require,module,exports){
"use strict";

var __extends = void 0 && (void 0).__extends || function () {
  var _extendStatics = function extendStatics(d, b) {
    _extendStatics = Object.setPrototypeOf || {
      __proto__: []
    } instanceof Array && function (d, b) {
      d.__proto__ = b;
    } || function (d, b) {
      for (var p in b) {
        if (b.hasOwnProperty(p)) d[p] = b[p];
      }
    };

    return _extendStatics(d, b);
  };

  return function (d, b) {
    _extendStatics(d, b);

    function __() {
      this.constructor = d;
    }

    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
  };
}();

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.getlineStringAttr = exports.getColor = exports.Element = exports.ReadXml = void 0;

var constant_1 = require("../common/constant");

var method_1 = require("../common/method");

var xmloperation =
/** @class */
function () {
  function xmloperation() {}
  /**
  * @param tag Search xml tag name , div,title etc.
  * @param file Xml string
  * @return Xml element string
  */


  xmloperation.prototype.getElementsByOneTag = function (tag, file) {
    //<a:[^/>: ]+?>.*?</a:[^/>: ]+?>
    var readTagReg;

    if (tag.indexOf("|") > -1) {
      var tags = tag.split("|"),
          tagsRegTxt = "";

      for (var i = 0; i < tags.length; i++) {
        var t = tags[i];
        tagsRegTxt += "|<" + t + " [^>]+?[^/]>[\\s\\S]*?</" + t + ">|<" + t + " [^>]+?/>|<" + t + ">[\\s\\S]*?</" + t + ">|<" + t + "/>";
      }

      tagsRegTxt = tagsRegTxt.substr(1, tagsRegTxt.length);
      readTagReg = new RegExp(tagsRegTxt, "g");
    } else {
      readTagReg = new RegExp("<" + tag + " [^>]+?[^/]>[\\s\\S]*?</" + tag + ">|<" + tag + " [^>]+?/>|<" + tag + ">[\\s\\S]*?</" + tag + ">|<" + tag + "/>", "g");
    }

    var ret = file.match(readTagReg);

    if (ret == null) {
      return [];
    } else {
      return ret;
    }
  };

  return xmloperation;
}();

var ReadXml =
/** @class */
function (_super) {
  __extends(ReadXml, _super);

  function ReadXml(files) {
    var _this = _super.call(this) || this;

    _this.originFile = files;
    return _this;
  }
  /**
  * @param path Search xml tag group , div,title etc.
  * @param fileName One of uploadfileList, uploadfileList is file group, {key:value}
  * @return Xml element calss
  */


  ReadXml.prototype.getElementsByTagName = function (path, fileName) {
    var file = this.getFileByName(fileName);
    var pathArr = path.split("/"),
        ret;

    for (var key in pathArr) {
      var path_1 = pathArr[key];

      if (ret == undefined) {
        ret = this.getElementsByOneTag(path_1, file);
      } else {
        if (ret instanceof Array) {
          var items = [];

          for (var key_1 in ret) {
            var item = ret[key_1];
            items = items.concat(this.getElementsByOneTag(path_1, item));
          }

          ret = items;
        } else {
          ret = this.getElementsByOneTag(path_1, ret);
        }
      }
    }

    var elements = [];

    for (var i = 0; i < ret.length; i++) {
      var ele = new Element(ret[i]);
      elements.push(ele);
    }

    return elements;
  };
  /**
  * @param name One of uploadfileList's name, search for file by this parameter
  * @retrun Select a file from uploadfileList
  */


  ReadXml.prototype.getFileByName = function (name) {
    for (var fileKey in this.originFile) {
      if (fileKey.indexOf(name) > -1) {
        return this.originFile[fileKey];
      }
    }

    return "";
  };

  return ReadXml;
}(xmloperation);

exports.ReadXml = ReadXml;

var Element =
/** @class */
function (_super) {
  __extends(Element, _super);

  function Element(str) {
    var _this = _super.call(this) || this;

    _this.elementString = str;

    _this.setValue();

    var readAttrReg = new RegExp('[a-zA-Z0-9_:]*?=".*?"', "g");

    var attrList = _this.container.match(readAttrReg);

    _this.attributeList = {};

    if (attrList != null) {
      for (var key in attrList) {
        var attrFull = attrList[key]; // let al= attrFull.split("=");

        if (attrFull.length == 0) {
          continue;
        }

        var attrKey = attrFull.substr(0, attrFull.indexOf('='));
        var attrValue = attrFull.substr(attrFull.indexOf('=') + 1);

        if (attrKey == null || attrValue == null || attrKey.length == 0 || attrValue.length == 0) {
          continue;
        }

        _this.attributeList[attrKey] = attrValue.substr(1, attrValue.length - 2);
      }
    }

    return _this;
  }
  /**
  * @param name Get attribute by key in element
  * @return Single attribute
  */


  Element.prototype.get = function (name) {
    return this.attributeList[name];
  };
  /**
  * @param tag Get elements by tag in elementString
  * @return Element group
  */


  Element.prototype.getInnerElements = function (tag) {
    var ret = this.getElementsByOneTag(tag, this.elementString);
    var elements = [];

    for (var i = 0; i < ret.length; i++) {
      var ele = new Element(ret[i]);
      elements.push(ele);
    }

    if (elements.length == 0) {
      return null;
    }

    return elements;
  };
  /**
  * @desc get xml dom value and container, <container>value</container>
  */


  Element.prototype.setValue = function () {
    var str = this.elementString;

    if (str.substr(str.length - 2, 2) == "/>") {
      this.value = "";
      this.container = str;
    } else {
      var firstTag = this.getFirstTag();
      var firstTagReg = new RegExp("(<" + firstTag + " [^>]+?[^/]>)([\\s\\S]*?)</" + firstTag + ">|(<" + firstTag + ">)([\\s\\S]*?)</" + firstTag + ">", "g");
      var result = firstTagReg.exec(str);

      if (result != null) {
        if (result[1] != null) {
          this.container = result[1];
          this.value = result[2];
        } else {
          this.container = result[3];
          this.value = result[4];
        }
      }
    }
  };
  /**
  * @desc get xml dom first tag, <a><b></b></a>, get a
  */


  Element.prototype.getFirstTag = function () {
    var str = this.elementString;
    var firstTag = str.substr(0, str.indexOf(' '));

    if (firstTag == "" || firstTag.indexOf(">") > -1) {
      firstTag = str.substr(0, str.indexOf('>'));
    }

    firstTag = firstTag.substr(1, firstTag.length);
    return firstTag;
  };

  return Element;
}(xmloperation);

exports.Element = Element;

function combineIndexedColor(indexedColorsInner, indexedColors) {
  var ret = {};

  if (indexedColorsInner == null || indexedColorsInner.length == 0) {
    return indexedColors;
  }

  for (var key in indexedColors) {
    var value = indexedColors[key],
        kn = parseInt(key);
    var inner = indexedColorsInner[kn];

    if (inner == null) {
      ret[key] = value;
    } else {
      var rgb = inner.attributeList.rgb;
      ret[key] = rgb;
    }
  }

  return ret;
} //clrScheme:Element[]


function getColor(color, styles, type) {
  if (type === void 0) {
    type = "g";
  }

  var attrList = color.attributeList;
  var clrScheme = styles["clrScheme"];
  var indexedColorsInner = styles["indexedColors"];
  var mruColorsInner = styles["mruColors"];
  var indexedColorsList = combineIndexedColor(indexedColorsInner, constant_1.indexedColors);
  var indexed = attrList.indexed,
      rgb = attrList.rgb,
      theme = attrList.theme,
      tint = attrList.tint;
  var bg;

  if (indexed != null) {
    var indexedNum = parseInt(indexed);
    bg = indexedColorsList[indexedNum];

    if (bg != null) {
      bg = bg.substring(bg.length - 6, bg.length);
      bg = "#" + bg;
    }
  } else if (rgb != null) {
    rgb = rgb.substring(rgb.length - 6, rgb.length);
    bg = "#" + rgb;
  } else if (theme != null) {
    var themeNum = parseInt(theme);

    if (themeNum == 0) {
      themeNum = 1;
    } else if (themeNum == 1) {
      themeNum = 0;
    } else if (themeNum == 2) {
      themeNum = 3;
    } else if (themeNum == 3) {
      themeNum = 2;
    }

    var clrSchemeElement = clrScheme[themeNum];

    if (clrSchemeElement != null) {
      var clrs = clrSchemeElement.getInnerElements("a:sysClr|a:srgbClr");

      if (clrs != null) {
        var clr = clrs[0];
        var clrAttrList = clr.attributeList; // console.log(clr.container, );

        if (clr.container.indexOf("sysClr") > -1) {
          // if(type=="g" && clrAttrList.val=="windowText"){
          //     bg = null;
          // }
          // else if((type=="t" || type=="b") && clrAttrList.val=="window"){
          //     bg = null;
          // }                    
          // else 
          if (clrAttrList.lastClr != null) {
            bg = "#" + clrAttrList.lastClr;
          } else if (clrAttrList.val != null) {
            bg = "#" + clrAttrList.val;
          }
        } else if (clr.container.indexOf("srgbClr") > -1) {
          // console.log(clrAttrList.val);
          bg = "#" + clrAttrList.val;
        }
      }
    }
  }

  if (tint != null) {
    var tintNum = parseFloat(tint);

    if (bg != null) {
      bg = method_1.LightenDarkenColor(bg, tintNum);
    }
  }

  return bg;
}

exports.getColor = getColor;
/**
 * @dom xml attribute object
 * @attr attribute name
 * @d if attribute is null, return default value
 * @return attribute value
*/

function getlineStringAttr(frpr, attr) {
  var attrEle = frpr.getInnerElements(attr),
      value;

  if (attrEle != null && attrEle.length > 0) {
    if (attr == "b" || attr == "i" || attr == "strike") {
      value = "1";
    } else if (attr == "u") {
      var v = attrEle[0].attributeList.val;

      if (v == "double") {
        value = "2";
      } else if (v == "singleAccounting") {
        value = "3";
      } else if (v == "doubleAccounting") {
        value = "4";
      } else {
        value = "1";
      }
    } else if (attr == "vertAlign") {
      var v = attrEle[0].attributeList.val;

      if (v == "subscript") {
        value = "1";
      } else if (v == "superscript") {
        value = "2";
      }
    } else {
      value = attrEle[0].attributeList.val;
    }
  }

  return value;
}

exports.getlineStringAttr = getlineStringAttr;

},{"../common/constant":16,"../common/method":18}],16:[function(require,module,exports){
"use strict";

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.DATA_VERIFICATION_TYPE2_MAP = exports.COMMON_TYPE2 = exports.DATA_VERIFICATION_MAP = exports.fontFamilys = exports.numFmtDefaultMap = exports.borderTypes = exports.OEM_CHARSET = exports.indexedColors = exports.numFmtDefault = exports.BuiltInCellStyles = exports.ST_CellType = exports.workbookRels = exports.theme1File = exports.worksheetFilePath = exports.sharedStringsFile = exports.stylesFile = exports.calcChainFile = exports.workBookFile = exports.contentTypesFile = exports.appFile = exports.coreFile = exports.columeHeader_word_index = exports.columeHeader_word = void 0;
exports.columeHeader_word = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
exports.columeHeader_word_index = {
  'A': 0,
  'B': 1,
  'C': 2,
  'D': 3,
  'E': 4,
  'F': 5,
  'G': 6,
  'H': 7,
  'I': 8,
  'J': 9,
  'K': 10,
  'L': 11,
  'M': 12,
  'N': 13,
  'O': 14,
  'P': 15,
  'Q': 16,
  'R': 17,
  'S': 18,
  'T': 19,
  'U': 20,
  'V': 21,
  'W': 22,
  'X': 23,
  'Y': 24,
  'Z': 25
};
exports.coreFile = "docProps/core.xml";
exports.appFile = "docProps/app.xml";
exports.contentTypesFile = "[Content_Types].xml";
exports.workBookFile = "xl/workbook.xml";
exports.calcChainFile = "xl/calcChain.xml";
exports.stylesFile = "xl/styles.xml";
exports.sharedStringsFile = "xl/sharedStrings.xml";
exports.worksheetFilePath = "xl/worksheets/";
exports.theme1File = "xl/theme/theme1.xml";
exports.workbookRels = "xl/_rels/workbook.xml.rels"; //Excel Built-In cell type

exports.ST_CellType = {
  "Boolean": "b",
  "Date": "d",
  "Error": "e",
  "InlineString": "inlineStr",
  "Number": "n",
  "SharedString": "s",
  "String": "str"
}; //Excel Built-In cell style

exports.BuiltInCellStyles = {
  "0": "Normal"
};
exports.numFmtDefault = {
  "0": 'General',
  "1": '0',
  "2": '0.00',
  "3": '#,##0',
  "4": '#,##0.00',
  "9": '0%',
  "10": '0.00%',
  "11": '0.00E+00',
  "12": '# ?/?',
  "13": '# ??/??',
  "14": 'm/d/yy',
  "15": 'd-mmm-yy',
  "16": 'd-mmm',
  "17": 'mmm-yy',
  "18": 'h:mm AM/PM',
  "19": 'h:mm:ss AM/PM',
  "20": 'h:mm',
  "21": 'h:mm:ss',
  "22": 'm/d/yy h:mm',
  "37": '#,##0 ;(#,##0)',
  "38": '#,##0 ;[Red](#,##0)',
  "39": '#,##0.00;(#,##0.00)',
  "40": '#,##0.00;[Red](#,##0.00)',
  "45": 'mm:ss',
  "46": '[h]:mm:ss',
  "47": 'mmss.0',
  "48": '##0.0E+0',
  "49": '@'
};
exports.indexedColors = {
  "0": '00000000',
  "1": '00FFFFFF',
  "2": '00FF0000',
  "3": '0000FF00',
  "4": '000000FF',
  "5": '00FFFF00',
  "6": '00FF00FF',
  "7": '0000FFFF',
  "8": '00000000',
  "9": '00FFFFFF',
  "10": '00FF0000',
  "11": '0000FF00',
  "12": '000000FF',
  "13": '00FFFF00',
  "14": '00FF00FF',
  "15": '0000FFFF',
  "16": '00800000',
  "17": '00008000',
  "18": '00000080',
  "19": '00808000',
  "20": '00800080',
  "21": '00008080',
  "22": '00C0C0C0',
  "23": '00808080',
  "24": '009999FF',
  "25": '00993366',
  "26": '00FFFFCC',
  "27": '00CCFFFF',
  "28": '00660066',
  "29": '00FF8080',
  "30": '000066CC',
  "31": '00CCCCFF',
  "32": '00000080',
  "33": '00FF00FF',
  "34": '00FFFF00',
  "35": '0000FFFF',
  "36": '00800080',
  "37": '00800000',
  "38": '00008080',
  "39": '000000FF',
  "40": '0000CCFF',
  "41": '00CCFFFF',
  "42": '00CCFFCC',
  "43": '00FFFF99',
  "44": '0099CCFF',
  "45": '00FF99CC',
  "46": '00CC99FF',
  "47": '00FFCC99',
  "48": '003366FF',
  "49": '0033CCCC',
  "50": '0099CC00',
  "51": '00FFCC00',
  "52": '00FF9900',
  "53": '00FF6600',
  "54": '00666699',
  "55": '00969696',
  "56": '00003366',
  "57": '00339966',
  "58": '00003300',
  "59": '00333300',
  "60": '00993300',
  "61": '00993366',
  "62": '00333399',
  "63": '00333333',
  "64": null,
  "65": null
};
exports.OEM_CHARSET = {
  "0": "ANSI_CHARSET",
  "1": "DEFAULT_CHARSET",
  "2": "SYMBOL_CHARSET",
  "77": "MAC_CHARSET",
  "128": "SHIFTJIS_CHARSET",
  "129": "HANGUL_CHARSET",
  "130": "JOHAB_CHARSET",
  "134": "GB2312_CHARSET",
  "136": "CHINESEBIG5_CHARSET",
  "161": "GREEK_CHARSET",
  "162": "TURKISH_CHARSET",
  "163": "VIETNAMESE_CHARSET",
  "177": "HEBREW_CHARSET",
  "178": "ARABIC_CHARSET",
  "186": "BALTIC_CHARSET",
  "204": "RUSSIAN_CHARSET",
  "222": "THAI_CHARSET",
  "238": "EASTEUROPE_CHARSET",
  "255": "OEM_CHARSET"
};
exports.borderTypes = {
  "none": 0,
  "thin": 1,
  "hair": 2,
  "dotted": 3,
  "dashed": 4,
  "dashDot": 5,
  "dashDotDot": 6,
  "double": 7,
  "medium": 8,
  "mediumDashed": 9,
  "mediumDashDot": 10,
  "mediumDashDotDot": 11,
  "slantDashDot": 12,
  "thick": 13
};
exports.numFmtDefaultMap = {
  "yyyy/m/d;@": "yyyy/MM/dd",
  "yyyy&quot;年&quot;m&quot;月&quot;d&quot;日&quot;;@": "yyyy&quot;年&quot;MM&quot;月&quot;dd&quot;日&quot;",
  "[$-409]yyyy/m/d\\ h:mm\\ AM/PM;@": "yyyy/MM/dd hh:mm AM/PM"
};
exports.fontFamilys = {
  "0": "defualt",
  "1": "Roman",
  "2": "Swiss",
  "3": "Modern",
  "4": "Script",
  "5": "Decorative"
};
exports.DATA_VERIFICATION_MAP = {
  list: "dropdown",
  whole: "number_integer",
  decimal: "number_decimal",
  custom: "text_content",
  textLength: "text_length",
  date: "date",
  "unknown1": "number",
  "unknown2": "checkbox",
  "unknown3": "validity"
};
exports.COMMON_TYPE2 = ["number", "number_integer", "number_decimal", "text_length"];
exports.DATA_VERIFICATION_TYPE2_MAP = {
  common: {
    between: "bw",
    notBetween: "nb",
    equal: "eq",
    notEqualTo: "ne",
    moreThanThe: "gt",
    lessThan: "lt",
    greaterOrEqualTo: "gte",
    lessThanOrEqualTo: "lte"
  },
  text_content: {
    include: "include",
    exclude: "exclude",
    equal: "equal"
  },
  date: {
    between: "bw",
    notBetween: "nb",
    equal: "eq",
    notEqualTo: "ne",
    earlierThan: "bf",
    noEarlierThan: "nbf",
    laterThan: "af",
    noLaterThan: "naf"
  },
  validity: {
    card: "card",
    phone: "phone"
  }
};

},{}],17:[function(require,module,exports){
"use strict";

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.ToContext2D = exports.FromEMF = exports.UDOC = void 0;
exports.UDOC = {};
exports.UDOC.G = {
  concat: function concat(p, r) {
    for (var i = 0; i < r.cmds.length; i++) {
      p.cmds.push(r.cmds[i]);
    }

    for (var i = 0; i < r.crds.length; i++) {
      p.crds.push(r.crds[i]);
    }
  },
  getBB: function getBB(ps) {
    var x0 = 1e99,
        y0 = 1e99,
        x1 = -x0,
        y1 = -y0;

    for (var i = 0; i < ps.length; i += 2) {
      var x = ps[i],
          y = ps[i + 1];
      if (x < x0) x0 = x;else if (x > x1) x1 = x;
      if (y < y0) y0 = y;else if (y > y1) y1 = y;
    }

    return [x0, y0, x1, y1];
  },
  rectToPath: function rectToPath(r) {
    return {
      cmds: ["M", "L", "L", "L", "Z"],
      crds: [r[0], r[1], r[2], r[1], r[2], r[3], r[0], r[3]]
    };
  },
  // a inside b
  insideBox: function insideBox(a, b) {
    return b[0] <= a[0] && b[1] <= a[1] && a[2] <= b[2] && a[3] <= b[3];
  },
  isBox: function isBox(p, bb) {
    var sameCrd8 = function sameCrd8(pcrd, crds) {
      for (var o = 0; o < 8; o += 2) {
        var eq = true;

        for (var j = 0; j < 8; j++) {
          if (Math.abs(crds[j] - pcrd[j + o & 7]) >= 2) {
            eq = false;
            break;
          }
        }

        if (eq) return true;
      }

      return false;
    };

    if (p.cmds.length > 10) return false;
    var cmds = p.cmds.join(""),
        crds = p.crds;
    var sameRect = false;

    if (cmds == "MLLLZ" && crds.length == 8 || cmds == "MLLLLZ" && crds.length == 10) {
      if (crds.length == 10) crds = crds.slice(0, 8);
      var x0 = bb[0],
          y0 = bb[1],
          x1 = bb[2],
          y1 = bb[3];
      if (!sameRect) sameRect = sameCrd8(crds, [x0, y0, x1, y0, x1, y1, x0, y1]);
      if (!sameRect) sameRect = sameCrd8(crds, [x0, y1, x1, y1, x1, y0, x0, y0]);
    }

    return sameRect;
  },
  boxArea: function boxArea(a) {
    var w = a[2] - a[0],
        h = a[3] - a[1];
    return w * h;
  },
  newPath: function newPath(gst) {
    gst.pth = {
      cmds: [],
      crds: []
    };
  },
  moveTo: function moveTo(gst, x, y) {
    var p = exports.UDOC.M.multPoint(gst.ctm, [x, y]); //if(gst.cpos[0]==p[0] && gst.cpos[1]==p[1]) return;

    gst.pth.cmds.push("M");
    gst.pth.crds.push(p[0], p[1]);
    gst.cpos = p;
  },
  lineTo: function lineTo(gst, x, y) {
    var p = exports.UDOC.M.multPoint(gst.ctm, [x, y]);
    if (gst.cpos[0] == p[0] && gst.cpos[1] == p[1]) return;
    gst.pth.cmds.push("L");
    gst.pth.crds.push(p[0], p[1]);
    gst.cpos = p;
  },
  curveTo: function curveTo(gst, x1, y1, x2, y2, x3, y3) {
    var p;
    p = exports.UDOC.M.multPoint(gst.ctm, [x1, y1]);
    x1 = p[0];
    y1 = p[1];
    p = exports.UDOC.M.multPoint(gst.ctm, [x2, y2]);
    x2 = p[0];
    y2 = p[1];
    p = exports.UDOC.M.multPoint(gst.ctm, [x3, y3]);
    x3 = p[0];
    y3 = p[1];
    gst.cpos = p;
    gst.pth.cmds.push("C");
    gst.pth.crds.push(x1, y1, x2, y2, x3, y3);
  },
  closePath: function closePath(gst) {
    gst.pth.cmds.push("Z");
  },
  arc: function arc(gst, x, y, r, a0, a1, neg) {
    // circle from a0 counter-clock-wise to a1
    if (neg) while (a1 > a0) {
      a1 -= 2 * Math.PI;
    } else while (a1 < a0) {
      a1 += 2 * Math.PI;
    }
    var th = (a1 - a0) / 4;
    var x0 = Math.cos(th / 2),
        y0 = -Math.sin(th / 2);
    var x1 = (4 - x0) / 3,
        y1 = y0 == 0 ? y0 : (1 - x0) * (3 - x0) / (3 * y0);
    var x2 = x1,
        y2 = -y1;
    var x3 = x0,
        y3 = -y0;
    var p0 = [x0, y0],
        p1 = [x1, y1],
        p2 = [x2, y2],
        p3 = [x3, y3];
    var pth = {
      cmds: [gst.pth.cmds.length == 0 ? "M" : "L", "C", "C", "C", "C"],
      crds: [x0, y0, x1, y1, x2, y2, x3, y3]
    };
    var rot = [1, 0, 0, 1, 0, 0];
    exports.UDOC.M.rotate(rot, -th);

    for (var i = 0; i < 3; i++) {
      p1 = exports.UDOC.M.multPoint(rot, p1);
      p2 = exports.UDOC.M.multPoint(rot, p2);
      p3 = exports.UDOC.M.multPoint(rot, p3);
      pth.crds.push(p1[0], p1[1], p2[0], p2[1], p3[0], p3[1]);
    }

    var sc = [r, 0, 0, r, x, y];
    exports.UDOC.M.rotate(rot, -a0 + th / 2);
    exports.UDOC.M.concat(rot, sc);
    exports.UDOC.M.multArray(rot, pth.crds);
    exports.UDOC.M.multArray(gst.ctm, pth.crds);
    exports.UDOC.G.concat(gst.pth, pth);
    var y = pth.crds.pop();
    x = pth.crds.pop();
    gst.cpos = [x, y];
  },
  toPoly: function toPoly(p) {
    if (p.cmds[0] != "M" || p.cmds[p.cmds.length - 1] != "Z") return null;

    for (var i = 1; i < p.cmds.length - 1; i++) {
      if (p.cmds[i] != "L") return null;
    }

    var out = [],
        cl = p.crds.length;
    if (p.crds[0] == p.crds[cl - 2] && p.crds[1] == p.crds[cl - 1]) cl -= 2;

    for (var i = 0; i < cl; i += 2) {
      out.push([p.crds[i], p.crds[i + 1]]);
    }

    if (exports.UDOC.G.polyArea(p.crds) < 0) out.reverse();
    return out;
  },
  fromPoly: function fromPoly(p) {
    var o = {
      cmds: [],
      crds: []
    };

    for (var i = 0; i < p.length; i++) {
      o.crds.push(p[i][0], p[i][1]);
      o.cmds.push(i == 0 ? "M" : "L");
    }

    o.cmds.push("Z");
    return o;
  },
  polyArea: function polyArea(p) {
    if (p.length < 6) return 0;
    var l = p.length - 2;
    var sum = (p[0] - p[l]) * (p[l + 1] + p[1]);

    for (var i = 0; i < l; i += 2) {
      sum += (p[i + 2] - p[i]) * (p[i + 1] + p[i + 3]);
    }

    return -sum * 0.5;
  },
  polyClip: function polyClip(p0, p1) {
    var cp1, cp2, s, e;

    var inside = function inside(p) {
      return (cp2[0] - cp1[0]) * (p[1] - cp1[1]) > (cp2[1] - cp1[1]) * (p[0] - cp1[0]);
    };

    var isc = function isc() {
      var dc = [cp1[0] - cp2[0], cp1[1] - cp2[1]],
          dp = [s[0] - e[0], s[1] - e[1]],
          n1 = cp1[0] * cp2[1] - cp1[1] * cp2[0],
          n2 = s[0] * e[1] - s[1] * e[0],
          n3 = 1.0 / (dc[0] * dp[1] - dc[1] * dp[0]);
      return [(n1 * dp[0] - n2 * dc[0]) * n3, (n1 * dp[1] - n2 * dc[1]) * n3];
    };

    var out = p0;
    cp1 = p1[p1.length - 1];

    for (var j in p1) {
      var cp2 = p1[j];
      var inp = out;
      out = [];
      s = inp[inp.length - 1]; //last on the input list

      for (var i in inp) {
        var e = inp[i];

        if (inside(e)) {
          if (!inside(s)) {
            out.push(isc());
          }

          out.push(e);
        } else if (inside(s)) {
          out.push(isc());
        }

        s = e;
      }

      cp1 = cp2;
    }

    return out;
  }
};
exports.UDOC.M = {
  getScale: function getScale(m) {
    return Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2]));
  },
  translate: function translate(m, x, y) {
    exports.UDOC.M.concat(m, [1, 0, 0, 1, x, y]);
  },
  rotate: function rotate(m, a) {
    exports.UDOC.M.concat(m, [Math.cos(a), -Math.sin(a), Math.sin(a), Math.cos(a), 0, 0]);
  },
  scale: function scale(m, x, y) {
    exports.UDOC.M.concat(m, [x, 0, 0, y, 0, 0]);
  },
  concat: function concat(m, w) {
    var a = m[0],
        b = m[1],
        c = m[2],
        d = m[3],
        tx = m[4],
        ty = m[5];
    m[0] = a * w[0] + b * w[2];
    m[1] = a * w[1] + b * w[3];
    m[2] = c * w[0] + d * w[2];
    m[3] = c * w[1] + d * w[3];
    m[4] = tx * w[0] + ty * w[2] + w[4];
    m[5] = tx * w[1] + ty * w[3] + w[5];
  },
  invert: function invert(m) {
    var a = m[0],
        b = m[1],
        c = m[2],
        d = m[3],
        tx = m[4],
        ty = m[5],
        adbc = a * d - b * c;
    m[0] = d / adbc;
    m[1] = -b / adbc;
    m[2] = -c / adbc;
    m[3] = a / adbc;
    m[4] = (c * ty - d * tx) / adbc;
    m[5] = (b * tx - a * ty) / adbc;
  },
  multPoint: function multPoint(m, p) {
    var x = p[0],
        y = p[1];
    return [x * m[0] + y * m[2] + m[4], x * m[1] + y * m[3] + m[5]];
  },
  multArray: function multArray(m, a) {
    for (var i = 0; i < a.length; i += 2) {
      var x = a[i],
          y = a[i + 1];
      a[i] = x * m[0] + y * m[2] + m[4];
      a[i + 1] = x * m[1] + y * m[3] + m[5];
    }
  }
};
exports.UDOC.C = {
  srgbGamma: function srgbGamma(x) {
    return x < 0.0031308 ? 12.92 * x : 1.055 * Math.pow(x, 1.0 / 2.4) - 0.055;
  },
  cmykToRgb: function cmykToRgb(clr) {
    var c = clr[0],
        m = clr[1],
        y = clr[2],
        k = clr[3]; // return [1-Math.min(1,c+k), 1-Math.min(1, m+k), 1-Math.min(1,y+k)];

    var r = 255 + c * (-4.387332384609988 * c + 54.48615194189176 * m + 18.82290502165302 * y + 212.25662451639585 * k + -285.2331026137004) + m * (1.7149763477362134 * m - 5.6096736904047315 * y + -17.873870861415444 * k - 5.497006427196366) + y * (-2.5217340131683033 * y - 21.248923337353073 * k + 17.5119270841813) + k * (-21.86122147463605 * k - 189.48180835922747);
    var g = 255 + c * (8.841041422036149 * c + 60.118027045597366 * m + 6.871425592049007 * y + 31.159100130055922 * k + -79.2970844816548) + m * (-15.310361306967817 * m + 17.575251261109482 * y + 131.35250912493976 * k - 190.9453302588951) + y * (4.444339102852739 * y + 9.8632861493405 * k - 24.86741582555878) + k * (-20.737325471181034 * k - 187.80453709719578);
    var b = 255 + c * (0.8842522430003296 * c + 8.078677503112928 * m + 30.89978309703729 * y - 0.23883238689178934 * k + -14.183576799673286) + m * (10.49593273432072 * m + 63.02378494754052 * y + 50.606957656360734 * k - 112.23884253719248) + y * (0.03296041114873217 * y + 115.60384449646641 * k + -193.58209356861505) + k * (-22.33816807309886 * k - 180.12613974708367);
    return [Math.max(0, Math.min(1, r / 255)), Math.max(0, Math.min(1, g / 255)), Math.max(0, Math.min(1, b / 255))]; //var iK = 1-c[3];  
    //return [(1-c[0])*iK, (1-c[1])*iK, (1-c[2])*iK];  
  },
  labToRgb: function labToRgb(lab) {
    var k = 903.3,
        e = 0.008856,
        L = lab[0],
        a = lab[1],
        b = lab[2];
    var fy = (L + 16) / 116,
        fy3 = fy * fy * fy;
    var fz = fy - b / 200,
        fz3 = fz * fz * fz;
    var fx = a / 500 + fy,
        fx3 = fx * fx * fx;
    var zr = fz3 > e ? fz3 : (116 * fz - 16) / k;
    var yr = fy3 > e ? fy3 : (116 * fy - 16) / k;
    var xr = fx3 > e ? fx3 : (116 * fx - 16) / k;
    var X = xr * 96.72,
        Y = yr * 100,
        Z = zr * 81.427,
        xyz = [X / 100, Y / 100, Z / 100];
    var x2s = [3.1338561, -1.6168667, -0.4906146, -0.9787684, 1.9161415, 0.0334540, 0.0719453, -0.2289914, 1.4052427];
    var rgb = [x2s[0] * xyz[0] + x2s[1] * xyz[1] + x2s[2] * xyz[2], x2s[3] * xyz[0] + x2s[4] * xyz[1] + x2s[5] * xyz[2], x2s[6] * xyz[0] + x2s[7] * xyz[1] + x2s[8] * xyz[2]];

    for (var i = 0; i < 3; i++) {
      rgb[i] = Math.max(0, Math.min(1, exports.UDOC.C.srgbGamma(rgb[i])));
    }

    return rgb;
  }
};

exports.UDOC.getState = function (crds) {
  return {
    font: exports.UDOC.getFont(),
    dd: {
      flat: 1
    },
    space: "/DeviceGray",
    // fill
    ca: 1,
    colr: [0, 0, 0],
    sspace: "/DeviceGray",
    // stroke
    CA: 1,
    COLR: [0, 0, 0],
    bmode: "/Normal",
    SA: false,
    OPM: 0,
    AIS: false,
    OP: false,
    op: false,
    SMask: "/None",
    lwidth: 1,
    lcap: 0,
    ljoin: 0,
    mlimit: 10,
    SM: 0.1,
    doff: 0,
    dash: [],
    ctm: [1, 0, 0, 1, 0, 0],
    cpos: [0, 0],
    pth: {
      cmds: [],
      crds: []
    },
    cpth: crds ? exports.UDOC.G.rectToPath(crds) : null // clipping path

  };
};

exports.UDOC.getFont = function () {
  return {
    Tc: 0,
    Tw: 0,
    Th: 100,
    Tl: 0,
    Tf: "Helvetica-Bold",
    Tfs: 1,
    Tmode: 0,
    Trise: 0,
    Tk: 0,
    Tal: 0,
    Tun: 0,
    Tm: [1, 0, 0, 1, 0, 0],
    Tlm: [1, 0, 0, 1, 0, 0],
    Trm: [1, 0, 0, 1, 0, 0]
  };
};

exports.FromEMF = function () {};

exports.FromEMF.Parse = function (buff, genv) {
  buff = new Uint8Array(buff);
  var off = 0; //console.log(buff.slice(0,32));

  var prms = {
    fill: false,
    strk: false,
    bb: [0, 0, 1, 1],
    wbb: [0, 0, 1, 1],
    fnt: {
      nam: "Arial",
      hgh: 25,
      und: false,
      orn: 0
    },
    tclr: [0, 0, 0],
    talg: 0
  },
      gst,
      tab = [],
      sts = [];
  var rI = exports.FromEMF.B.readShort,
      rU = exports.FromEMF.B.readUshort,
      rI32 = exports.FromEMF.B.readInt,
      rU32 = exports.FromEMF.B.readUint,
      rF32 = exports.FromEMF.B.readFloat;
  var opn = 0;

  while (true) {
    var fnc = rU32(buff, off);
    off += 4;
    var fnm = exports.FromEMF.K[fnc];
    var siz = rU32(buff, off);
    off += 4; //if(gst && isNaN(gst.ctm[0])) throw "e";
    //console.log(fnc,fnm,siz);

    var loff = off; //if(opn++==253) break;

    var obj = null,
        oid = 0; //console.log(fnm, siz);

    if (false) {} else if (fnm == "EOF") {
      break;
    } else if (fnm == "HEADER") {
      prms.bb = exports.FromEMF._readBox(buff, loff);
      loff += 16; //console.log(fnm, prms.bb);

      genv.StartPage(prms.bb[0], prms.bb[1], prms.bb[2], prms.bb[3]);
      gst = exports.UDOC.getState(prms.bb);
    } else if (fnm == "SAVEDC") sts.push(JSON.stringify(gst), JSON.stringify(prms));else if (fnm == "RESTOREDC") {
      var dif = rI32(buff, loff);
      loff += 4;

      while (dif < -1) {
        sts.pop();
        sts.pop();
      }

      prms = JSON.parse(sts.pop());
      gst = JSON.parse(sts.pop());
    } else if (fnm == "SELECTCLIPPATH") {
      gst.cpth = JSON.parse(JSON.stringify(gst.pth));
    } else if (["SETMAPMODE", "SETPOLYFILLMODE", "SETBKMODE"
    /*,"SETVIEWPORTEXTEX"*/
    , "SETICMMODE", "SETROP2", "EXTSELECTCLIPRGN"].indexOf(fnm) != -1) {} //else if(fnm=="INTERSECTCLIPRECT") {  var r=prms.crct=FromEMF._readBox(buff, loff);  /*var y0=r[1],y1=r[3]; if(y0>y1){r[1]=y1; r[3]=y0;}*/ console.log(prms.crct);  }
    else if (fnm == "SETMITERLIMIT") gst.mlimit = rU32(buff, loff);else if (fnm == "SETTEXTCOLOR") prms.tclr = [buff[loff] / 255, buff[loff + 1] / 255, buff[loff + 2] / 255];else if (fnm == "SETTEXTALIGN") prms.talg = rU32(buff, loff);else if (fnm == "SETVIEWPORTEXTEX" || fnm == "SETVIEWPORTORGEX") {
      if (prms.vbb == null) prms.vbb = [];
      var coff = fnm == "SETVIEWPORTORGEX" ? 0 : 2;
      prms.vbb[coff] = rI32(buff, loff);
      loff += 4;
      prms.vbb[coff + 1] = rI32(buff, loff);
      loff += 4; //console.log(prms.vbb);

      if (fnm == "SETVIEWPORTEXTEX") exports.FromEMF._updateCtm(prms, gst);
    } else if (fnm == "SETWINDOWEXTEX" || fnm == "SETWINDOWORGEX") {
      var coff = fnm == "SETWINDOWORGEX" ? 0 : 2;
      prms.wbb[coff] = rI32(buff, loff);
      loff += 4;
      prms.wbb[coff + 1] = rI32(buff, loff);
      loff += 4;
      if (fnm == "SETWINDOWEXTEX") exports.FromEMF._updateCtm(prms, gst);
    } //else if(fnm=="SETMETARGN") {}
    else if (fnm == "COMMENT") {
      var ds = rU32(buff, loff);
      loff += 4;
    } else if (fnm == "SELECTOBJECT") {
      var ind = rU32(buff, loff);
      loff += 4; //console.log(ind.toString(16), tab, tab[ind]);

      if (ind == 0x80000000) {
        prms.fill = true;
        gst.colr = [1, 1, 1];
      } // white brush
      else if (ind == 0x80000005) {
        prms.fill = false;
      } // null brush
      else if (ind == 0x80000007) {
        prms.strk = true;
        prms.lwidth = 1;
        gst.COLR = [0, 0, 0];
      } // black pen
      else if (ind == 0x80000008) {
        prms.strk = false;
      } // null  pen
      else if (ind == 0x8000000d) {} // system font
      else if (ind == 0x8000000e) {} // device default font
      else {
        var co = tab[ind]; //console.log(ind, co);

        if (co.t == "b") {
          prms.fill = co.stl != 1;

          if (co.stl == 0) {} else if (co.stl == 1) {} else throw co.stl + " e";

          gst.colr = co.clr;
        } else if (co.t == "p") {
          prms.strk = co.stl != 5;
          gst.lwidth = co.wid;
          gst.COLR = co.clr;
        } else if (co.t == "f") {
          prms.fnt = co;
          gst.font.Tf = co.nam;
          gst.font.Tfs = Math.abs(co.hgh);
          gst.font.Tun = co.und;
        } else throw "e";
      }
    } else if (fnm == "DELETEOBJECT") {
      var ind = rU32(buff, loff);
      loff += 4;
      if (tab[ind] != null) tab[ind] = null;else throw "e";
    } else if (fnm == "CREATEBRUSHINDIRECT") {
      oid = rU32(buff, loff);
      loff += 4;
      obj = {
        t: "b"
      };
      obj.stl = rU32(buff, loff);
      loff += 4;
      obj.clr = [buff[loff] / 255, buff[loff + 1] / 255, buff[loff + 2] / 255];
      loff += 4;
      obj.htc = rU32(buff, loff);
      loff += 4; //console.log(oid, obj);
    } else if (fnm == "CREATEPEN" || fnm == "EXTCREATEPEN") {
      oid = rU32(buff, loff);
      loff += 4;
      obj = {
        t: "p"
      };

      if (fnm == "EXTCREATEPEN") {
        loff += 16;
        obj.stl = rU32(buff, loff);
        loff += 4;
        obj.wid = rU32(buff, loff);
        loff += 4; //obj.stl = rU32(buff, loff);  

        loff += 4;
      } else {
        obj.stl = rU32(buff, loff);
        loff += 4;
        obj.wid = rU32(buff, loff);
        loff += 4;
        loff += 4;
      }

      obj.clr = [buff[loff] / 255, buff[loff + 1] / 255, buff[loff + 2] / 255];
      loff += 4;
    } else if (fnm == "EXTCREATEFONTINDIRECTW") {
      oid = rU32(buff, loff);
      loff += 4;
      obj = {
        t: "f",
        nam: ""
      };
      obj.hgh = rI32(buff, loff);
      loff += 4;
      loff += 4 * 2;
      obj.orn = rI32(buff, loff) / 10;
      loff += 4;
      var wgh = rU32(buff, loff);
      loff += 4; //console.log(fnm, obj.orn, wgh);
      //console.log(rU32(buff,loff), rU32(buff,loff+4), buff.slice(loff,loff+8));

      obj.und = buff[loff + 1];
      obj.stk = buff[loff + 2];
      loff += 4 * 2;

      while (rU(buff, loff) != 0) {
        obj.nam += String.fromCharCode(rU(buff, loff));
        loff += 2;
      }

      if (wgh > 500) obj.nam += "-Bold"; //console.log(wgh, obj.nam);
    } else if (fnm == "EXTTEXTOUTW") {
      //console.log(buff.slice(loff-8, loff-8+siz));
      loff += 16;
      var mod = rU32(buff, loff);
      loff += 4; //console.log(mod);

      var scx = rF32(buff, loff);
      loff += 4;
      var scy = rF32(buff, loff);
      loff += 4;
      var rfx = rI32(buff, loff);
      loff += 4;
      var rfy = rI32(buff, loff);
      loff += 4; //console.log(mod, scx, scy,rfx,rfy);

      gst.font.Tm = [1, 0, 0, -1, 0, 0];
      exports.UDOC.M.rotate(gst.font.Tm, prms.fnt.orn * Math.PI / 180);
      exports.UDOC.M.translate(gst.font.Tm, rfx, rfy);
      var alg = prms.talg; //console.log(alg.toString(2));

      if ((alg & 6) == 6) gst.font.Tal = 2;else if ((alg & 7) == 0) gst.font.Tal = 0;else throw alg + " e";

      if ((alg & 24) == 24) {} // baseline
      else if ((alg & 24) == 0) exports.UDOC.M.translate(gst.font.Tm, 0, gst.font.Tfs);else throw "e";

      var crs = rU32(buff, loff);
      loff += 4;
      var ofs = rU32(buff, loff);
      loff += 4;
      var ops = rU32(buff, loff);
      loff += 4; //if(ops!=0) throw "e";
      //console.log(ofs,ops,crs);

      loff += 16;
      var ofD = rU32(buff, loff);
      loff += 4; //console.log(ops, ofD, loff, ofs+off-8);

      ofs += off - 8; //console.log(crs, ops);

      var str = "";

      for (var i = 0; i < crs; i++) {
        var cc = rU(buff, ofs + i * 2);
        str += String.fromCharCode(cc);
      }

      ;
      var oclr = gst.colr;
      gst.colr = prms.tclr; //console.log(str, gst.colr, gst.font.Tm);
      //var otfs = gst.font.Tfs;  gst.font.Tfs *= 1/gst.ctm[0];

      genv.PutText(gst, str, str.length * gst.font.Tfs * 0.5);
      gst.colr = oclr; //gst.font.Tfs = otfs;
      //console.log(rfx, rfy, scx, ops, rcX, rcY, rcW, rcH, offDx, str);
    } else if (fnm == "BEGINPATH") {
      exports.UDOC.G.newPath(gst);
    } else if (fnm == "ENDPATH") {} else if (fnm == "CLOSEFIGURE") exports.UDOC.G.closePath(gst);else if (fnm == "MOVETOEX") {
      exports.UDOC.G.moveTo(gst, rI32(buff, loff), rI32(buff, loff + 4));
    } else if (fnm == "LINETO") {
      if (gst.pth.cmds.length == 0) {
        var im = gst.ctm.slice(0);
        exports.UDOC.M.invert(im);
        var p = exports.UDOC.M.multPoint(im, gst.cpos);
        exports.UDOC.G.moveTo(gst, p[0], p[1]);
      }

      exports.UDOC.G.lineTo(gst, rI32(buff, loff), rI32(buff, loff + 4));
    } else if (fnm == "POLYGON" || fnm == "POLYGON16" || fnm == "POLYLINE" || fnm == "POLYLINE16" || fnm == "POLYLINETO" || fnm == "POLYLINETO16") {
      loff += 16;
      var ndf = fnm.startsWith("POLYGON"),
          isTo = fnm.indexOf("TO") != -1;
      var cnt = rU32(buff, loff);
      loff += 4;
      if (!isTo) exports.UDOC.G.newPath(gst);
      loff = exports.FromEMF._drawPoly(buff, loff, cnt, gst, fnm.endsWith("16") ? 2 : 4, ndf, isTo);
      if (!isTo) exports.FromEMF._draw(genv, gst, prms, ndf); //console.log(prms, gst.lwidth);
      //console.log(JSON.parse(JSON.stringify(gst.pth)));
    } else if (fnm == "POLYPOLYGON16") {
      loff += 16;
      var ndf = fnm.startsWith("POLYPOLYGON"),
          isTo = fnm.indexOf("TO") != -1;
      var nop = rU32(buff, loff);
      loff += 4;
      loff += 4;
      var pi = loff;
      loff += nop * 4;
      if (!isTo) exports.UDOC.G.newPath(gst);

      for (var i = 0; i < nop; i++) {
        var ppp = rU(buff, pi + i * 4);
        loff = exports.FromEMF._drawPoly(buff, loff, ppp, gst, fnm.endsWith("16") ? 2 : 4, ndf, isTo);
      }

      if (!isTo) exports.FromEMF._draw(genv, gst, prms, ndf);
    } else if (fnm == "POLYBEZIER" || fnm == "POLYBEZIER16" || fnm == "POLYBEZIERTO" || fnm == "POLYBEZIERTO16") {
      loff += 16;
      var is16 = fnm.endsWith("16"),
          rC = is16 ? rI : rI32,
          nl = is16 ? 2 : 4;
      var cnt = rU32(buff, loff);
      loff += 4;

      if (fnm.indexOf("TO") == -1) {
        exports.UDOC.G.moveTo(gst, rC(buff, loff), rC(buff, loff + nl));
        loff += 2 * nl;
        cnt--;
      }

      while (cnt > 0) {
        exports.UDOC.G.curveTo(gst, rC(buff, loff), rC(buff, loff + nl), rC(buff, loff + 2 * nl), rC(buff, loff + 3 * nl), rC(buff, loff + 4 * nl), rC(buff, loff + 5 * nl));
        loff += 6 * nl;
        cnt -= 3;
      } //console.log(JSON.parse(JSON.stringify(gst.pth)));

    } else if (fnm == "RECTANGLE" || fnm == "ELLIPSE") {
      exports.UDOC.G.newPath(gst);

      var bx = exports.FromEMF._readBox(buff, loff);

      if (fnm == "RECTANGLE") {
        exports.UDOC.G.moveTo(gst, bx[0], bx[1]);
        exports.UDOC.G.lineTo(gst, bx[2], bx[1]);
        exports.UDOC.G.lineTo(gst, bx[2], bx[3]);
        exports.UDOC.G.lineTo(gst, bx[0], bx[3]);
      } else {
        var x = (bx[0] + bx[2]) / 2,
            y = (bx[1] + bx[3]) / 2;
        exports.UDOC.G.arc(gst, x, y, (bx[2] - bx[0]) / 2, 0, 2 * Math.PI, false);
      }

      exports.UDOC.G.closePath(gst);

      exports.FromEMF._draw(genv, gst, prms, true); //console.log(prms, gst.lwidth);

    } else if (fnm == "FILLPATH") genv.Fill(gst, false);else if (fnm == "STROKEPATH") genv.Stroke(gst);else if (fnm == "STROKEANDFILLPATH") {
      genv.Fill(gst, false);
      genv.Stroke(gst);
    } else if (fnm == "SETWORLDTRANSFORM" || fnm == "MODIFYWORLDTRANSFORM") {
      var mat = [];

      for (var i = 0; i < 6; i++) {
        mat.push(rF32(buff, loff + i * 4));
      }

      loff += 24; //console.log(fnm, gst.ctm.slice(0), mat);

      if (fnm == "SETWORLDTRANSFORM") gst.ctm = mat;else {
        var mod = rU32(buff, loff);
        loff += 4;

        if (mod == 2) {
          var om = gst.ctm;
          gst.ctm = mat;
          exports.UDOC.M.concat(gst.ctm, om);
        } else throw "e";
      }
    } else if (fnm == "SETSTRETCHBLTMODE") {
      var sm = rU32(buff, loff);
      loff += 4;
    } else if (fnm == "STRETCHDIBITS") {
      var bx = exports.FromEMF._readBox(buff, loff);

      loff += 16;
      var xD = rI32(buff, loff);
      loff += 4;
      var yD = rI32(buff, loff);
      loff += 4;
      var xS = rI32(buff, loff);
      loff += 4;
      var yS = rI32(buff, loff);
      loff += 4;
      var wS = rI32(buff, loff);
      loff += 4;
      var hS = rI32(buff, loff);
      loff += 4;
      var ofH = rU32(buff, loff) + off - 8;
      loff += 4;
      var szH = rU32(buff, loff);
      loff += 4;
      var ofB = rU32(buff, loff) + off - 8;
      loff += 4;
      var szB = rU32(buff, loff);
      loff += 4;
      var usg = rU32(buff, loff);
      loff += 4;
      if (usg != 0) throw "e";
      var bop = rU32(buff, loff);
      loff += 4;
      var wD = rI32(buff, loff);
      loff += 4;
      var hD = rI32(buff, loff);
      loff += 4; //console.log(bop, wD, hD);
      //console.log(ofH, szH, ofB, szB, ofH+40);
      //console.log(bx, xD,yD,wD,hD);
      //console.log(xS,yS,wS,hS);
      //console.log(ofH,szH,ofB,szB,usg,bop);

      var hl = rU32(buff, ofH);
      ofH += 4;
      var w = rU32(buff, ofH);
      ofH += 4;
      var h = rU32(buff, ofH);
      ofH += 4;
      if (w != wS || h != hS) throw "e";
      var ps = rU(buff, ofH);
      ofH += 2;
      var bc = rU(buff, ofH);
      ofH += 2;
      if (bc != 8 && bc != 24 && bc != 32) throw bc + " e";
      var cpr = rU32(buff, ofH);
      ofH += 4;
      if (cpr != 0) throw cpr + " e";
      var sz = rU32(buff, ofH);
      ofH += 4;
      var xpm = rU32(buff, ofH);
      ofH += 4;
      var ypm = rU32(buff, ofH);
      ofH += 4;
      var cu = rU32(buff, ofH);
      ofH += 4;
      var ci = rU32(buff, ofH);
      ofH += 4; //console.log(hl, w, h, ps, bc, cpr, sz, xpm, ypm, cu, ci);
      //console.log(hl,w,h,",",xS,yS,wS,hS,",",xD,yD,wD,hD,",",xpm,ypm);

      var rl = Math.floor((w * ps * bc + 31 & ~31) / 8);
      var img = new Uint8Array(w * h * 4);

      if (bc == 8) {
        for (var y = 0; y < h; y++) {
          for (var x = 0; x < w; x++) {
            var qi = y * w + x << 2,
                ind = buff[ofB + (h - 1 - y) * rl + x] << 2;
            img[qi] = buff[ofH + ind + 2];
            img[qi + 1] = buff[ofH + ind + 1];
            img[qi + 2] = buff[ofH + ind + 0];
            img[qi + 3] = 255;
          }
        }
      }

      if (bc == 24) {
        for (var y = 0; y < h; y++) {
          for (var x = 0; x < w; x++) {
            var qi = y * w + x << 2,
                ti = ofB + (h - 1 - y) * rl + x * 3;
            img[qi] = buff[ti + 2];
            img[qi + 1] = buff[ti + 1];
            img[qi + 2] = buff[ti + 0];
            img[qi + 3] = 255;
          }
        }
      }

      if (bc == 32) {
        for (var y = 0; y < h; y++) {
          for (var x = 0; x < w; x++) {
            var qi = y * w + x << 2,
                ti = ofB + (h - 1 - y) * rl + x * 4;
            img[qi] = buff[ti + 2];
            img[qi + 1] = buff[ti + 1];
            img[qi + 2] = buff[ti + 0];
            img[qi + 3] = buff[ti + 3];
          }
        }
      }

      var ctm = gst.ctm.slice(0);
      gst.ctm = [1, 0, 0, 1, 0, 0];
      exports.UDOC.M.scale(gst.ctm, wD, -hD);
      exports.UDOC.M.translate(gst.ctm, xD, yD + hD);
      exports.UDOC.M.concat(gst.ctm, ctm);
      genv.PutImage(gst, img, w, h);
      gst.ctm = ctm;
    } else {
      console.log(fnm, siz);
    }

    if (obj != null) tab[oid] = obj;
    off += siz - 8;
  } //genv.Stroke(gst);


  genv.ShowPage();
  genv.Done();
};

exports.FromEMF._readBox = function (buff, off) {
  var b = [];

  for (var i = 0; i < 4; i++) {
    b[i] = exports.FromEMF.B.readInt(buff, off + i * 4);
  }

  return b;
};

exports.FromEMF._updateCtm = function (prms, gst) {
  var mat = [1, 0, 0, 1, 0, 0];
  var wbb = prms.wbb,
      bb = prms.bb,
      vbb = prms.vbb && prms.vbb.length == 4 ? prms.vbb : prms.bb; //var y0 = bb[1], y1 = bb[3];  bb[1]=Math.min(y0,y1);  bb[3]=Math.max(y0,y1);

  exports.UDOC.M.translate(mat, -wbb[0], -wbb[1]);
  exports.UDOC.M.scale(mat, 1 / wbb[2], 1 / wbb[3]);
  exports.UDOC.M.scale(mat, vbb[2], vbb[3]); //UDOC.M.scale(mat, vbb[2]/(bb[2]-bb[0]), vbb[3]/(bb[3]-bb[1]));
  //UDOC.M.scale(mat, bb[2]-bb[0],bb[3]-bb[1]);

  gst.ctm = mat;
};

exports.FromEMF._draw = function (genv, gst, prms, needFill) {
  if (prms.fill && needFill) genv.Fill(gst, false);
  if (prms.strk && gst.lwidth != 0) genv.Stroke(gst);
};

exports.FromEMF._drawPoly = function (buff, off, ppp, gst, nl, clos, justLine) {
  var rS = nl == 2 ? exports.FromEMF.B.readShort : exports.FromEMF.B.readInt;

  for (var j = 0; j < ppp; j++) {
    var px = rS(buff, off);
    off += nl;
    var py = rS(buff, off);
    off += nl;
    if (j == 0 && !justLine) exports.UDOC.G.moveTo(gst, px, py);else exports.UDOC.G.lineTo(gst, px, py);
  }

  if (clos) exports.UDOC.G.closePath(gst);
  return off;
};

exports.FromEMF.B = {
  uint8: new Uint8Array(4),
  readShort: function readShort(buff, p) {
    var u8 = exports.FromEMF.B.uint8;
    u8[0] = buff[p];
    u8[1] = buff[p + 1];
    return exports.FromEMF.B.int16[0];
  },
  readUshort: function readUshort(buff, p) {
    var u8 = exports.FromEMF.B.uint8;
    u8[0] = buff[p];
    u8[1] = buff[p + 1];
    return exports.FromEMF.B.uint16[0];
  },
  readInt: function readInt(buff, p) {
    var u8 = exports.FromEMF.B.uint8;
    u8[0] = buff[p];
    u8[1] = buff[p + 1];
    u8[2] = buff[p + 2];
    u8[3] = buff[p + 3];
    return exports.FromEMF.B.int32[0];
  },
  readUint: function readUint(buff, p) {
    var u8 = exports.FromEMF.B.uint8;
    u8[0] = buff[p];
    u8[1] = buff[p + 1];
    u8[2] = buff[p + 2];
    u8[3] = buff[p + 3];
    return exports.FromEMF.B.uint32[0];
  },
  readFloat: function readFloat(buff, p) {
    var u8 = exports.FromEMF.B.uint8;
    u8[0] = buff[p];
    u8[1] = buff[p + 1];
    u8[2] = buff[p + 2];
    u8[3] = buff[p + 3];
    return exports.FromEMF.B.flot32[0];
  },
  readASCII: function readASCII(buff, p, l) {
    var s = "";

    for (var i = 0; i < l; i++) {
      s += String.fromCharCode(buff[p + i]);
    }

    return s;
  }
};
exports.FromEMF.B.int16 = new Int16Array(exports.FromEMF.B.uint8.buffer);
exports.FromEMF.B.uint16 = new Uint16Array(exports.FromEMF.B.uint8.buffer);
exports.FromEMF.B.int32 = new Int32Array(exports.FromEMF.B.uint8.buffer);
exports.FromEMF.B.uint32 = new Uint32Array(exports.FromEMF.B.uint8.buffer);
exports.FromEMF.B.flot32 = new Float32Array(exports.FromEMF.B.uint8.buffer);
exports.FromEMF.C = {
  EMR_HEADER: 0x00000001,
  EMR_POLYBEZIER: 0x00000002,
  EMR_POLYGON: 0x00000003,
  EMR_POLYLINE: 0x00000004,
  EMR_POLYBEZIERTO: 0x00000005,
  EMR_POLYLINETO: 0x00000006,
  EMR_POLYPOLYLINE: 0x00000007,
  EMR_POLYPOLYGON: 0x00000008,
  EMR_SETWINDOWEXTEX: 0x00000009,
  EMR_SETWINDOWORGEX: 0x0000000A,
  EMR_SETVIEWPORTEXTEX: 0x0000000B,
  EMR_SETVIEWPORTORGEX: 0x0000000C,
  EMR_SETBRUSHORGEX: 0x0000000D,
  EMR_EOF: 0x0000000E,
  EMR_SETPIXELV: 0x0000000F,
  EMR_SETMAPPERFLAGS: 0x00000010,
  EMR_SETMAPMODE: 0x00000011,
  EMR_SETBKMODE: 0x00000012,
  EMR_SETPOLYFILLMODE: 0x00000013,
  EMR_SETROP2: 0x00000014,
  EMR_SETSTRETCHBLTMODE: 0x00000015,
  EMR_SETTEXTALIGN: 0x00000016,
  EMR_SETCOLORADJUSTMENT: 0x00000017,
  EMR_SETTEXTCOLOR: 0x00000018,
  EMR_SETBKCOLOR: 0x00000019,
  EMR_OFFSETCLIPRGN: 0x0000001A,
  EMR_MOVETOEX: 0x0000001B,
  EMR_SETMETARGN: 0x0000001C,
  EMR_EXCLUDECLIPRECT: 0x0000001D,
  EMR_INTERSECTCLIPRECT: 0x0000001E,
  EMR_SCALEVIEWPORTEXTEX: 0x0000001F,
  EMR_SCALEWINDOWEXTEX: 0x00000020,
  EMR_SAVEDC: 0x00000021,
  EMR_RESTOREDC: 0x00000022,
  EMR_SETWORLDTRANSFORM: 0x00000023,
  EMR_MODIFYWORLDTRANSFORM: 0x00000024,
  EMR_SELECTOBJECT: 0x00000025,
  EMR_CREATEPEN: 0x00000026,
  EMR_CREATEBRUSHINDIRECT: 0x00000027,
  EMR_DELETEOBJECT: 0x00000028,
  EMR_ANGLEARC: 0x00000029,
  EMR_ELLIPSE: 0x0000002A,
  EMR_RECTANGLE: 0x0000002B,
  EMR_ROUNDRECT: 0x0000002C,
  EMR_ARC: 0x0000002D,
  EMR_CHORD: 0x0000002E,
  EMR_PIE: 0x0000002F,
  EMR_SELECTPALETTE: 0x00000030,
  EMR_CREATEPALETTE: 0x00000031,
  EMR_SETPALETTEENTRIES: 0x00000032,
  EMR_RESIZEPALETTE: 0x00000033,
  EMR_REALIZEPALETTE: 0x00000034,
  EMR_EXTFLOODFILL: 0x00000035,
  EMR_LINETO: 0x00000036,
  EMR_ARCTO: 0x00000037,
  EMR_POLYDRAW: 0x00000038,
  EMR_SETARCDIRECTION: 0x00000039,
  EMR_SETMITERLIMIT: 0x0000003A,
  EMR_BEGINPATH: 0x0000003B,
  EMR_ENDPATH: 0x0000003C,
  EMR_CLOSEFIGURE: 0x0000003D,
  EMR_FILLPATH: 0x0000003E,
  EMR_STROKEANDFILLPATH: 0x0000003F,
  EMR_STROKEPATH: 0x00000040,
  EMR_FLATTENPATH: 0x00000041,
  EMR_WIDENPATH: 0x00000042,
  EMR_SELECTCLIPPATH: 0x00000043,
  EMR_ABORTPATH: 0x00000044,
  EMR_COMMENT: 0x00000046,
  EMR_FILLRGN: 0x00000047,
  EMR_FRAMERGN: 0x00000048,
  EMR_INVERTRGN: 0x00000049,
  EMR_PAINTRGN: 0x0000004A,
  EMR_EXTSELECTCLIPRGN: 0x0000004B,
  EMR_BITBLT: 0x0000004C,
  EMR_STRETCHBLT: 0x0000004D,
  EMR_MASKBLT: 0x0000004E,
  EMR_PLGBLT: 0x0000004F,
  EMR_SETDIBITSTODEVICE: 0x00000050,
  EMR_STRETCHDIBITS: 0x00000051,
  EMR_EXTCREATEFONTINDIRECTW: 0x00000052,
  EMR_EXTTEXTOUTA: 0x00000053,
  EMR_EXTTEXTOUTW: 0x00000054,
  EMR_POLYBEZIER16: 0x00000055,
  EMR_POLYGON16: 0x00000056,
  EMR_POLYLINE16: 0x00000057,
  EMR_POLYBEZIERTO16: 0x00000058,
  EMR_POLYLINETO16: 0x00000059,
  EMR_POLYPOLYLINE16: 0x0000005A,
  EMR_POLYPOLYGON16: 0x0000005B,
  EMR_POLYDRAW16: 0x0000005C,
  EMR_CREATEMONOBRUSH: 0x0000005D,
  EMR_CREATEDIBPATTERNBRUSHPT: 0x0000005E,
  EMR_EXTCREATEPEN: 0x0000005F,
  EMR_POLYTEXTOUTA: 0x00000060,
  EMR_POLYTEXTOUTW: 0x00000061,
  EMR_SETICMMODE: 0x00000062,
  EMR_CREATECOLORSPACE: 0x00000063,
  EMR_SETCOLORSPACE: 0x00000064,
  EMR_DELETECOLORSPACE: 0x00000065,
  EMR_GLSRECORD: 0x00000066,
  EMR_GLSBOUNDEDRECORD: 0x00000067,
  EMR_PIXELFORMAT: 0x00000068,
  EMR_DRAWESCAPE: 0x00000069,
  EMR_EXTESCAPE: 0x0000006A,
  EMR_SMALLTEXTOUT: 0x0000006C,
  EMR_FORCEUFIMAPPING: 0x0000006D,
  EMR_NAMEDESCAPE: 0x0000006E,
  EMR_COLORCORRECTPALETTE: 0x0000006F,
  EMR_SETICMPROFILEA: 0x00000070,
  EMR_SETICMPROFILEW: 0x00000071,
  EMR_ALPHABLEND: 0x00000072,
  EMR_SETLAYOUT: 0x00000073,
  EMR_TRANSPARENTBLT: 0x00000074,
  EMR_GRADIENTFILL: 0x00000076,
  EMR_SETLINKEDUFIS: 0x00000077,
  EMR_SETTEXTJUSTIFICATION: 0x00000078,
  EMR_COLORMATCHTOTARGETW: 0x00000079,
  EMR_CREATECOLORSPACEW: 0x0000007A
};
exports.FromEMF.K = []; // (function() {
//     var inp, out, stt;
//     inp = FromEMF.C;   out = FromEMF.K;   stt=4;
//     for(var p in inp) out[inp[p]] = p.slice(stt);
// }  )();

exports.ToContext2D = function (needPage, scale) {
  this.canvas = document.createElement("canvas");
  this.ctx = this.canvas.getContext("2d");
  this.bb = null;
  this.currPage = 0;
  this.needPage = needPage;
  this.scale = scale;
};

exports.ToContext2D.prototype.StartPage = function (x, y, w, h) {
  if (this.currPage != this.needPage) return;
  this.bb = [x, y, w, h];
  var scl = this.scale,
      dpr = window.devicePixelRatio;
  var cnv = this.canvas,
      ctx = this.ctx;
  cnv.width = Math.round(w * scl);
  cnv.height = Math.round(h * scl);
  ctx.translate(0, h * scl);
  ctx.scale(scl, -scl);
  cnv.setAttribute("style", "border:1px solid; width:" + cnv.width / dpr + "px; height:" + cnv.height / dpr + "px");
};

exports.ToContext2D.prototype.Fill = function (gst, evenOdd) {
  if (this.currPage != this.needPage) return;
  var ctx = this.ctx;
  ctx.beginPath();

  this._setStyle(gst, ctx);

  this._draw(gst.pth, ctx);

  ctx.fill();
};

exports.ToContext2D.prototype.Stroke = function (gst) {
  if (this.currPage != this.needPage) return;
  var ctx = this.ctx;
  ctx.beginPath();

  this._setStyle(gst, ctx);

  this._draw(gst.pth, ctx);

  ctx.stroke();
};

exports.ToContext2D.prototype.PutText = function (gst, str, stw) {
  if (this.currPage != this.needPage) return;

  var scl = this._scale(gst.ctm);

  var ctx = this.ctx;

  this._setStyle(gst, ctx);

  ctx.save();
  var m = [1, 0, 0, -1, 0, 0];

  this._concat(m, gst.font.Tm);

  this._concat(m, gst.ctm); //console.log(str, m, gst);  throw "e";


  ctx.transform(m[0], m[1], m[2], m[3], m[4], m[5]);
  ctx.fillText(str, 0, 0);
  ctx.restore();
};

exports.ToContext2D.prototype.PutImage = function (gst, buff, w, h, msk) {
  if (this.currPage != this.needPage) return;
  var ctx = this.ctx;

  if (buff.length == w * h * 4) {
    buff = buff.slice(0);
    if (msk && msk.length == w * h * 4) for (var i = 0; i < buff.length; i += 4) {
      buff[i + 3] = msk[i + 1];
    }
    var cnv = document.createElement("canvas"),
        cctx = cnv.getContext("2d");
    cnv.width = w;
    cnv.height = h;
    var imgd = cctx.createImageData(w, h);

    for (var i = 0; i < buff.length; i++) {
      imgd.data[i] = buff[i];
    }

    cctx.putImageData(imgd, 0, 0);
    ctx.save();
    var m = [1, 0, 0, 1, 0, 0];

    this._concat(m, [1 / w, 0, 0, -1 / h, 0, 1]);

    this._concat(m, gst.ctm);

    ctx.transform(m[0], m[1], m[2], m[3], m[4], m[5]);
    ctx.drawImage(cnv, 0, 0);
    ctx.restore();
  }
};

exports.ToContext2D.prototype.ShowPage = function () {
  this.currPage++;
};

exports.ToContext2D.prototype.Done = function () {};

function _flt(n) {
  return "" + parseFloat(n.toFixed(2));
}

exports.ToContext2D.prototype._setStyle = function (gst, ctx) {
  var scl = this._scale(gst.ctm);

  ctx.fillStyle = this._getFill(gst.colr, gst.ca, ctx);
  ctx.strokeStyle = this._getFill(gst.COLR, gst.CA, ctx);
  ctx.lineCap = ["butt", "round", "square"][gst.lcap];
  ctx.lineJoin = ["miter", "round", "bevel"][gst.ljoin];
  ctx.lineWidth = gst.lwidth * scl;
  var dsh = gst.dash.slice(0);

  for (var i = 0; i < dsh.length; i++) {
    dsh[i] = _flt(dsh[i] * scl);
  }

  ctx.setLineDash(dsh);
  ctx.miterLimit = gst.mlimit * scl;
  var fn = gst.font.Tf,
      ln = fn.toLowerCase();
  var p0 = ln.indexOf("bold") != -1 ? "bold " : "";
  var p1 = ln.indexOf("italic") != -1 || ln.indexOf("oblique") != -1 ? "italic " : "";
  ctx.font = p0 + p1 + gst.font.Tfs + "px \"" + fn + "\"";
};

exports.ToContext2D.prototype._getFill = function (colr, ca, ctx) {
  if (colr.typ == null) return this._colr(colr, ca);else {
    var grd = colr,
        crd = grd.crds,
        mat = grd.mat,
        scl = this._scale(mat),
        gf;

    if (grd.typ == "lin") {
      var p0 = this._multPoint(mat, crd.slice(0, 2)),
          p1 = this._multPoint(mat, crd.slice(2));

      gf = ctx.createLinearGradient(p0[0], p0[1], p1[0], p1[1]);
    } else if (grd.typ == "rad") {
      var p0 = this._multPoint(mat, crd.slice(0, 2)),
          p1 = this._multPoint(mat, crd.slice(3));

      gf = ctx.createRadialGradient(p0[0], p0[1], crd[2] * scl, p1[0], p1[1], crd[5] * scl);
    }

    for (var i = 0; i < grd.grad.length; i++) {
      gf.addColorStop(grd.grad[i][0], this._colr(grd.grad[i][1], ca));
    }

    return gf;
  }
};

exports.ToContext2D.prototype._colr = function (c, a) {
  return "rgba(" + Math.round(c[0] * 255) + "," + Math.round(c[1] * 255) + "," + Math.round(c[2] * 255) + "," + a + ")";
};

exports.ToContext2D.prototype._scale = function (m) {
  return Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2]));
};

exports.ToContext2D.prototype._concat = function (m, w) {
  var a = m[0],
      b = m[1],
      c = m[2],
      d = m[3],
      tx = m[4],
      ty = m[5];
  m[0] = a * w[0] + b * w[2];
  m[1] = a * w[1] + b * w[3];
  m[2] = c * w[0] + d * w[2];
  m[3] = c * w[1] + d * w[3];
  m[4] = tx * w[0] + ty * w[2] + w[4];
  m[5] = tx * w[1] + ty * w[3] + w[5];
};

exports.ToContext2D.prototype._multPoint = function (m, p) {
  var x = p[0],
      y = p[1];
  return [x * m[0] + y * m[2] + m[4], x * m[1] + y * m[3] + m[5]];
}, exports.ToContext2D.prototype._draw = function (path, ctx) {
  var c = 0,
      crds = path.crds;

  for (var j = 0; j < path.cmds.length; j++) {
    var cmd = path.cmds[j];

    if (cmd == "M") {
      ctx.moveTo(crds[c], crds[c + 1]);
      c += 2;
    } else if (cmd == "L") {
      ctx.lineTo(crds[c], crds[c + 1]);
      c += 2;
    } else if (cmd == "C") {
      ctx.bezierCurveTo(crds[c], crds[c + 1], crds[c + 2], crds[c + 3], crds[c + 4], crds[c + 5]);
      c += 6;
    } else if (cmd == "Q") {
      ctx.quadraticCurveTo(crds[c], crds[c + 1], crds[c + 2], crds[c + 3]);
      c += 4;
    } else if (cmd == "Z") {
      ctx.closePath();
    }
  }
};

},{}],18:[function(require,module,exports){
"use strict";

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.getMultiFormulaValue = exports.getPeelOffX14 = exports.getTransR1C1ToSequence = exports.getSingleSequenceToNum = exports.getSqrefRawArrFormat = exports.getRegionSequence = exports.getMultiSequenceToNum = exports.getBinaryContent = exports.isContainMultiType = exports.isKoera = exports.isJapanese = exports.isChinese = exports.fromulaRef = exports.escapeCharacter = exports.generateRandomIndex = exports.LightenDarkenColor = exports.getRowHeightPixel = exports.getColumnWidthPixel = exports.getXmlAttibute = exports.getPxByEMUs = exports.getptToPxRatioByDPI = exports.getcellrange = exports.getRangetxt = void 0;

var constant_1 = require("./constant");

function getRangetxt(range, sheettxt) {
  var row0 = range["row"][0],
      row1 = range["row"][1];
  var column0 = range["column"][0],
      column1 = range["column"][1];

  if (row0 == null && row1 == null) {
    return sheettxt + chatatABC(column0) + ":" + chatatABC(column1);
  } else if (column0 == null && column1 == null) {
    return sheettxt + (row0 + 1) + ":" + (row1 + 1);
  } else {
    if (column0 == column1 && row0 == row1) {
      return sheettxt + chatatABC(column0) + (row0 + 1);
    } else {
      return sheettxt + chatatABC(column0) + (row0 + 1) + ":" + chatatABC(column1) + (row1 + 1);
    }
  }
}

exports.getRangetxt = getRangetxt;

function getcellrange(txt, sheets, sheetId) {
  if (sheets === void 0) {
    sheets = {};
  }

  if (sheetId === void 0) {
    sheetId = "1";
  }

  var val = txt.split("!");
  var sheettxt = "",
      rangetxt = "",
      sheetIndex = -1;

  if (val.length > 1) {
    sheettxt = val[0];
    rangetxt = val[1];
    var si = sheets[sheettxt];

    if (si == null) {
      sheetIndex = parseInt(sheetId);
    } else {
      sheetIndex = parseInt(si);
    }
  } else {
    sheetIndex = parseInt(sheetId);
    rangetxt = val[0];
  }

  if (rangetxt.indexOf(":") == -1) {
    var row = parseInt(rangetxt.replace(/[^0-9]/g, "")) - 1;
    var col = ABCatNum(rangetxt.replace(/[^A-Za-z]/g, ""));

    if (!isNaN(row) && !isNaN(col)) {
      return {
        "row": [row, row],
        "column": [col, col],
        "sheetIndex": sheetIndex
      };
    } else {
      return null;
    }
  } else {
    var rangetxtArray = rangetxt.split(":");
    var row = [],
        col = [];
    row[0] = parseInt(rangetxtArray[0].replace(/[^0-9]/g, "")) - 1;
    row[1] = parseInt(rangetxtArray[1].replace(/[^0-9]/g, "")) - 1; // if (isNaN(row[0])) {
    //     row[0] = 0;
    // }
    // if (isNaN(row[1])) {
    //     row[1] = sheetdata.length - 1;
    // }

    if (row[0] > row[1]) {
      return null;
    }

    col[0] = ABCatNum(rangetxtArray[0].replace(/[^A-Za-z]/g, ""));
    col[1] = ABCatNum(rangetxtArray[1].replace(/[^A-Za-z]/g, "")); // if (isNaN(col[0])) {
    //     col[0] = 0;
    // }
    // if (isNaN(col[1])) {
    //     col[1] = sheetdata[0].length - 1;
    // }

    if (col[0] > col[1]) {
      return null;
    }

    return {
      "row": row,
      "column": col,
      "sheetIndex": sheetIndex
    };
  }
}

exports.getcellrange = getcellrange; //列下标  字母转数字

function ABCatNum(abc) {
  abc = abc.toUpperCase();
  var abc_len = abc.length;

  if (abc_len == 0) {
    return NaN;
  }

  var abc_array = abc.split("");
  var wordlen = constant_1.columeHeader_word.length;
  var ret = 0;

  for (var i = abc_len - 1; i >= 0; i--) {
    if (i == abc_len - 1) {
      ret += constant_1.columeHeader_word_index[abc_array[i]];
    } else {
      ret += Math.pow(wordlen, abc_len - i - 1) * (constant_1.columeHeader_word_index[abc_array[i]] + 1);
    }
  }

  return ret;
} //列下标  数字转字母


function chatatABC(index) {
  var wordlen = constant_1.columeHeader_word.length;

  if (index < wordlen) {
    return constant_1.columeHeader_word[index];
  } else {
    var last = 0,
        pre = 0,
        ret = "";
    var i = 1,
        n = 0;

    while (index >= wordlen / (wordlen - 1) * (Math.pow(wordlen, i++) - 1)) {
      n = i;
    }

    var index_ab = index - wordlen / (wordlen - 1) * (Math.pow(wordlen, n - 1) - 1); //970

    last = index_ab + 1;

    for (var x = n; x > 0; x--) {
      var last1 = last,
          x1 = x; //-702=268, 3

      if (x == 1) {
        last1 = last1 % wordlen;

        if (last1 == 0) {
          last1 = 26;
        }

        return ret + constant_1.columeHeader_word[last1 - 1];
      }

      last1 = Math.ceil(last1 / Math.pow(wordlen, x - 1)); //last1 = last1 % wordlen;

      ret += constant_1.columeHeader_word[last1 - 1];

      if (x > 1) {
        last = last - (last1 - 1) * wordlen;
      }
    }
  }
}
/**
 * @return ratio, default 0.75 1in = 2.54cm = 25.4mm = 72pt = 6pc,  pt = 1/72 In, px = 1/dpi In
*/


function getptToPxRatioByDPI() {
  return 72 / 96;
}

exports.getptToPxRatioByDPI = getptToPxRatioByDPI;
/**
 * @emus EMUs, Excel drawing unit
 * @return pixel
*/

function getPxByEMUs(emus) {
  if (emus == null) {
    return 0;
  }

  var inch = emus / 914400;
  var pt = inch * 72;
  var px = pt / getptToPxRatioByDPI();
  return px;
}

exports.getPxByEMUs = getPxByEMUs;
/**
 * @dom xml attribute object
 * @attr attribute name
 * @d if attribute is null, return default value
 * @return attribute value
*/

function getXmlAttibute(dom, attr, d) {
  var value = dom[attr];
  value = value == null ? d : value;
  return value;
}

exports.getXmlAttibute = getXmlAttibute;
/**
 * @columnWidth Excel column width
 * @return pixel column width
*/

function getColumnWidthPixel(columnWidth) {
  var pix = Math.round((columnWidth - 0.83) * 8 + 5);
  return pix;
}

exports.getColumnWidthPixel = getColumnWidthPixel;
/**
 * @rowHeight Excel row height
 * @return pixel row height
*/

function getRowHeightPixel(rowHeight) {
  var pix = Math.round(rowHeight / getptToPxRatioByDPI());
  return pix;
}

exports.getRowHeightPixel = getRowHeightPixel;

function LightenDarkenColor(sixColor, tint) {
  var hex = sixColor.substring(sixColor.length - 6, sixColor.length);
  var rgbArray = hexToRgbArray("#" + hex);
  var hslArray = rgbToHsl(rgbArray[0], rgbArray[1], rgbArray[2]);

  if (tint > 0) {
    hslArray[2] = hslArray[2] * (1.0 - tint) + tint;
  } else if (tint < 0) {
    hslArray[2] = hslArray[2] * (1.0 + tint);
  } else {
    return "#" + hex;
  }

  var newRgbArray = hslToRgb(hslArray[0], hslArray[1], hslArray[2]);
  return rgbToHex("RGB(" + newRgbArray.join(",") + ")");
}

exports.LightenDarkenColor = LightenDarkenColor;

function rgbToHex(rgb) {
  //十六进制颜色值的正则表达式
  var reg = /^#([0-9a-fA-f]{3}|[0-9a-fA-f]{6})$/; // 如果是rgb颜色表示

  if (/^(rgb|RGB)/.test(rgb)) {
    var aColor = rgb.replace(/(?:\(|\)|rgb|RGB)*/g, "").split(",");
    var strHex = "#";

    for (var i = 0; i < aColor.length; i++) {
      var hex = Number(aColor[i]).toString(16);

      if (hex.length < 2) {
        hex = '0' + hex;
      }

      strHex += hex;
    }

    if (strHex.length !== 7) {
      strHex = rgb;
    }

    return strHex;
  } else if (reg.test(rgb)) {
    var aNum = rgb.replace(/#/, "").split("");

    if (aNum.length === 6) {
      return rgb;
    } else if (aNum.length === 3) {
      var numHex = "#";

      for (var i = 0; i < aNum.length; i += 1) {
        numHex += aNum[i] + aNum[i];
      }

      return numHex;
    }
  }

  return rgb;
}

function hexToRgb(hex) {
  var sColor = hex.toLowerCase(); //十六进制颜色值的正则表达式

  var reg = /^#([0-9a-fA-f]{3}|[0-9a-fA-f]{6})$/; // 如果是16进制颜色

  if (sColor && reg.test(sColor)) {
    if (sColor.length === 4) {
      var sColorNew = "#";

      for (var i = 1; i < 4; i += 1) {
        sColorNew += sColor.slice(i, i + 1).concat(sColor.slice(i, i + 1));
      }

      sColor = sColorNew;
    } //处理六位的颜色值


    var sColorChange = [];

    for (var i = 1; i < 7; i += 2) {
      sColorChange.push(parseInt("0x" + sColor.slice(i, i + 2)));
    }

    return "RGB(" + sColorChange.join(",") + ")";
  }

  return sColor;
}

function hexToRgbArray(hex) {
  var sColor = hex.toLowerCase(); //十六进制颜色值的正则表达式

  var reg = /^#([0-9a-fA-f]{3}|[0-9a-fA-f]{6})$/; // 如果是16进制颜色

  if (sColor && reg.test(sColor)) {
    if (sColor.length === 4) {
      var sColorNew = "#";

      for (var i = 1; i < 4; i += 1) {
        sColorNew += sColor.slice(i, i + 1).concat(sColor.slice(i, i + 1));
      }

      sColor = sColorNew;
    } //处理六位的颜色值


    var sColorChange = [];

    for (var i = 1; i < 7; i += 2) {
      sColorChange.push(parseInt("0x" + sColor.slice(i, i + 2)));
    }

    return sColorChange;
  }

  return null;
}
/**
 * HSL颜色值转换为RGB.
 * 换算公式改编自 http://en.wikipedia.org/wiki/HSL_color_space.
 * h, s, 和 l 设定在 [0, 1] 之间
 * 返回的 r, g, 和 b 在 [0, 255]之间
 *
 * @param   Number  h       色相
 * @param   Number  s       饱和度
 * @param   Number  l       亮度
 * @return  Array           RGB色值数值
 */


function hslToRgb(h, s, l) {
  var r, g, b;

  if (s == 0) {
    r = g = b = l; // achromatic
  } else {
    var hue2rgb = function hue2rgb(p, q, t) {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
      return p;
    };

    var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    var p = 2 * l - q;
    r = hue2rgb(p, q, h + 1 / 3);
    g = hue2rgb(p, q, h);
    b = hue2rgb(p, q, h - 1 / 3);
  }

  return [Math.round(r * 255), Math.round(g * 255), Math.round(b * 255)];
}
/**
 * RGB 颜色值转换为 HSL.
 * 转换公式参考自 http://en.wikipedia.org/wiki/HSL_color_space.
 * r, g, 和 b 需要在 [0, 255] 范围内
 * 返回的 h, s, 和 l 在 [0, 1] 之间
 *
 * @param   Number  r       红色色值
 * @param   Number  g       绿色色值
 * @param   Number  b       蓝色色值
 * @return  Array           HSL各值数组
 */


function rgbToHsl(r, g, b) {
  r /= 255, g /= 255, b /= 255;
  var max = Math.max(r, g, b),
      min = Math.min(r, g, b);
  var h,
      s,
      l = (max + min) / 2;

  if (max == min) {
    h = s = 0; // achromatic
  } else {
    var d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);

    switch (max) {
      case r:
        h = (g - b) / d + (g < b ? 6 : 0);
        break;

      case g:
        h = (b - r) / d + 2;
        break;

      case b:
        h = (r - g) / d + 4;
        break;
    }

    h /= 6;
  }

  return [h, s, l];
}

function generateRandomIndex(prefix) {
  if (prefix == null) {
    prefix = "Sheet";
  }

  var userAgent = window.navigator.userAgent.replace(/[^a-zA-Z0-9]/g, "").split("");
  var mid = "";

  for (var i = 0; i < 5; i++) {
    mid += userAgent[Math.round(Math.random() * (userAgent.length - 1))];
  }

  var time = new Date().getTime();
  return prefix + "_" + mid + "_" + time;
}

exports.generateRandomIndex = generateRandomIndex;

function escapeCharacter(str) {
  if (str == null || str.length == 0) {
    return str;
  }

  return str.replace(/&amp;/g, "&").replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ').replace(/&apos;/g, "'").replace(/&iexcl;/g, "¡").replace(/&cent;/g, "¢").replace(/&pound;/g, "£").replace(/&curren;/g, "¤").replace(/&yen;/g, "¥").replace(/&brvbar;/g, "¦").replace(/&sect;/g, "§").replace(/&uml;/g, "¨").replace(/&copy;/g, "©").replace(/&ordf;/g, "ª").replace(/&laquo;/g, "«").replace(/&not;/g, "¬").replace(/&shy;/g, "­").replace(/&reg;/g, "®").replace(/&macr;/g, "¯").replace(/&deg;/g, "°").replace(/&plusmn;/g, "±").replace(/&sup2;/g, "²").replace(/&sup3;/g, "³").replace(/&acute;/g, "´").replace(/&micro;/g, "µ").replace(/&para;/g, "¶").replace(/&middot;/g, "·").replace(/&cedil;/g, "¸").replace(/&sup1;/g, "¹").replace(/&ordm;/g, "º").replace(/&raquo;/g, "»").replace(/&frac14;/g, "¼").replace(/&frac12;/g, "½").replace(/&frac34;/g, "¾").replace(/&iquest;/g, "¿").replace(/&times;/g, "×").replace(/&divide;/g, "÷").replace(/&Agrave;/g, "À").replace(/&Aacute;/g, "Á").replace(/&Acirc;/g, "Â").replace(/&Atilde;/g, "Ã").replace(/&Auml;/g, "Ä").replace(/&Aring;/g, "Å").replace(/&AElig;/g, "Æ").replace(/&Ccedil;/g, "Ç").replace(/&Egrave;/g, "È").replace(/&Eacute;/g, "É").replace(/&Ecirc;/g, "Ê").replace(/&Euml;/g, "Ë").replace(/&Igrave;/g, "Ì").replace(/&Iacute;/g, "Í").replace(/&Icirc;/g, "Î").replace(/&Iuml;/g, "Ï").replace(/&ETH;/g, "Ð").replace(/&Ntilde;/g, "Ñ").replace(/&Ograve;/g, "Ò").replace(/&Oacute;/g, "Ó").replace(/&Ocirc;/g, "Ô").replace(/&Otilde;/g, "Õ").replace(/&Ouml;/g, "Ö").replace(/&Oslash;/g, "Ø").replace(/&Ugrave;/g, "Ù").replace(/&Uacute;/g, "Ú").replace(/&Ucirc;/g, "Û").replace(/&Uuml;/g, "Ü").replace(/&Yacute;/g, "Ý").replace(/&THORN;/g, "Þ").replace(/&szlig;/g, "ß").replace(/&agrave;/g, "à").replace(/&aacute;/g, "á").replace(/&acirc;/g, "â").replace(/&atilde;/g, "ã").replace(/&auml;/g, "ä").replace(/&aring;/g, "å").replace(/&aelig;/g, "æ").replace(/&ccedil;/g, "ç").replace(/&egrave;/g, "è").replace(/&eacute;/g, "é").replace(/&ecirc;/g, "ê").replace(/&euml;/g, "ë").replace(/&igrave;/g, "ì").replace(/&iacute;/g, "í").replace(/&icirc;/g, "î").replace(/&iuml;/g, "ï").replace(/&eth;/g, "ð").replace(/&ntilde;/g, "ñ").replace(/&ograve;/g, "ò").replace(/&oacute;/g, "ó").replace(/&ocirc;/g, "ô").replace(/&otilde;/g, "õ").replace(/&ouml;/g, "ö").replace(/&oslash;/g, "ø").replace(/&ugrave;/g, "ù").replace(/&uacute;/g, "ú").replace(/&ucirc;/g, "û").replace(/&uuml;/g, "ü").replace(/&yacute;/g, "ý").replace(/&thorn;/g, "þ").replace(/&yuml;/g, "ÿ");
}

exports.escapeCharacter = escapeCharacter;

var fromulaRef =
/** @class */
function () {
  function fromulaRef() {}

  fromulaRef.trim = function (str) {
    if (str == null) {
      str = "";
    }

    return str.replace(/(^\s*)|(\s*$)/g, "");
  };

  fromulaRef.functionCopy = function (txt, mode, step) {
    var _this = this;

    if (_this.operatorjson == null) {
      var arr = _this.operator.split("|"),
          op = {};

      for (var i_1 = 0; i_1 < arr.length; i_1++) {
        op[arr[i_1].toString()] = 1;
      }

      _this.operatorjson = op;
    }

    if (mode == null) {
      mode = "down";
    }

    if (step == null) {
      step = 1;
    }

    if (txt.substr(0, 1) == "=") {
      txt = txt.substr(1);
    }

    var funcstack = txt.split("");
    var i = 0,
        str = "",
        function_str = "",
        ispassby = true;
    var matchConfig = {
      "bracket": 0,
      "comma": 0,
      "squote": 0,
      "dquote": 0
    };

    while (i < funcstack.length) {
      var s = funcstack[i];

      if (s == "(" && matchConfig.dquote == 0) {
        matchConfig.bracket += 1;

        if (str.length > 0) {
          function_str += str + "(";
        } else {
          function_str += "(";
        }

        str = "";
      } else if (s == ")" && matchConfig.dquote == 0) {
        matchConfig.bracket -= 1;
        function_str += _this.functionCopy(str, mode, step) + ")";
        str = "";
      } else if (s == '"' && matchConfig.squote == 0) {
        if (matchConfig.dquote > 0) {
          function_str += str + '"';
          matchConfig.dquote -= 1;
          str = "";
        } else {
          matchConfig.dquote += 1;
          str += '"';
        }
      } else if (s == ',' && matchConfig.dquote == 0) {
        function_str += _this.functionCopy(str, mode, step) + ',';
        str = "";
      } else if (s == '&' && matchConfig.dquote == 0) {
        if (str.length > 0) {
          function_str += _this.functionCopy(str, mode, step) + "&";
          str = "";
        } else {
          function_str += "&";
        }
      } else if (s in _this.operatorjson && matchConfig.dquote == 0) {
        var s_next = "";

        if (i + 1 < funcstack.length) {
          s_next = funcstack[i + 1];
        }

        var p = i - 1,
            s_pre = null;

        if (p >= 0) {
          do {
            s_pre = funcstack[p--];
          } while (p >= 0 && s_pre == " ");
        }

        if (s + s_next in _this.operatorjson) {
          if (str.length > 0) {
            function_str += _this.functionCopy(str, mode, step) + s + s_next;
            str = "";
          } else {
            function_str += s + s_next;
          }

          i++;
        } else if (!/[^0-9]/.test(s_next) && s == "-" && (s_pre == "(" || s_pre == null || s_pre == "," || s_pre == " " || s_pre in _this.operatorjson)) {
          str += s;
        } else {
          if (str.length > 0) {
            function_str += _this.functionCopy(str, mode, step) + s;
            str = "";
          } else {
            function_str += s;
          }
        }
      } else {
        str += s;
      }

      if (i == funcstack.length - 1) {
        if (_this.iscelldata(_this.trim(str))) {
          if (mode == "down") {
            function_str += _this.downparam(_this.trim(str), step);
          } else if (mode == "up") {
            function_str += _this.upparam(_this.trim(str), step);
          } else if (mode == "left") {
            function_str += _this.leftparam(_this.trim(str), step);
          } else if (mode == "right") {
            function_str += _this.rightparam(_this.trim(str), step);
          }
        } else {
          function_str += _this.trim(str);
        }
      }

      i++;
    }

    return function_str;
  };

  fromulaRef.downparam = function (txt, step) {
    return this.updateparam("d", txt, step);
  };

  fromulaRef.upparam = function (txt, step) {
    return this.updateparam("u", txt, step);
  };

  fromulaRef.leftparam = function (txt, step) {
    return this.updateparam("l", txt, step);
  };

  fromulaRef.rightparam = function (txt, step) {
    return this.updateparam("r", txt, step);
  };

  fromulaRef.updateparam = function (orient, txt, step) {
    var _this = this;

    var val = txt.split("!"),
        rangetxt,
        prefix = "";

    if (val.length > 1) {
      rangetxt = val[1];
      prefix = val[0] + "!";
    } else {
      rangetxt = val[0];
    }

    if (rangetxt.indexOf(":") == -1) {
      var row = parseInt(rangetxt.replace(/[^0-9]/g, ""));
      var col = ABCatNum(rangetxt.replace(/[^A-Za-z]/g, ""));

      var freezonFuc = _this.isfreezonFuc(rangetxt);

      var $row = freezonFuc[0] ? "$" : "",
          $col = freezonFuc[1] ? "$" : "";

      if (orient == "u" && !freezonFuc[0]) {
        row -= step;
      } else if (orient == "r" && !freezonFuc[1]) {
        col += step;
      } else if (orient == "l" && !freezonFuc[1]) {
        col -= step;
      } else if (!freezonFuc[0]) {
        row += step;
      }

      if (row < 0 || col < 0) {
        return _this.error.r;
      }

      if (!isNaN(row) && !isNaN(col)) {
        return prefix + $col + chatatABC(col) + $row + row;
      } else if (!isNaN(row)) {
        return prefix + $row + row;
      } else if (!isNaN(col)) {
        return prefix + $col + chatatABC(col);
      } else {
        return txt;
      }
    } else {
      rangetxt = rangetxt.split(":");
      var row = [],
          col = [];
      row[0] = parseInt(rangetxt[0].replace(/[^0-9]/g, ""));
      row[1] = parseInt(rangetxt[1].replace(/[^0-9]/g, ""));

      if (row[0] > row[1]) {
        return txt;
      }

      col[0] = ABCatNum(rangetxt[0].replace(/[^A-Za-z]/g, ""));
      col[1] = ABCatNum(rangetxt[1].replace(/[^A-Za-z]/g, ""));

      if (col[0] > col[1]) {
        return txt;
      }

      var freezonFuc0 = _this.isfreezonFuc(rangetxt[0]);

      var freezonFuc1 = _this.isfreezonFuc(rangetxt[1]);

      var $row0 = freezonFuc0[0] ? "$" : "",
          $col0 = freezonFuc0[1] ? "$" : "";
      var $row1 = freezonFuc1[0] ? "$" : "",
          $col1 = freezonFuc1[1] ? "$" : "";

      if (orient == "u") {
        if (!freezonFuc0[0]) {
          row[0] -= step;
        }

        if (!freezonFuc1[0]) {
          row[1] -= step;
        }
      } else if (orient == "r") {
        if (!freezonFuc0[1]) {
          col[0] += step;
        }

        if (!freezonFuc1[1]) {
          col[1] += step;
        }
      } else if (orient == "l") {
        if (!freezonFuc0[1]) {
          col[0] -= step;
        }

        if (!freezonFuc1[1]) {
          col[1] -= step;
        }
      } else {
        if (!freezonFuc0[0]) {
          row[0] += step;
        }

        if (!freezonFuc1[0]) {
          row[1] += step;
        }
      }

      if (row[0] < 0 || col[0] < 0) {
        return _this.error.r;
      }

      if (isNaN(col[0]) && isNaN(col[1])) {
        return prefix + $row0 + row[0] + ":" + $row1 + row[1];
      } else if (isNaN(row[0]) && isNaN(row[1])) {
        return prefix + $col0 + chatatABC(col[0]) + ":" + $col1 + chatatABC(col[1]);
      } else {
        return prefix + $col0 + chatatABC(col[0]) + $row0 + row[0] + ":" + $col1 + chatatABC(col[1]) + $row1 + row[1];
      }
    }
  };

  fromulaRef.iscelldata = function (txt) {
    var val = txt.split("!"),
        rangetxt;

    if (val.length > 1) {
      rangetxt = val[1];
    } else {
      rangetxt = val[0];
    }

    var reg_cell = /^(([a-zA-Z]+)|([$][a-zA-Z]+))(([0-9]+)|([$][0-9]+))$/g; //增加正则判断单元格为字母+数字的格式：如 A1:B3

    var reg_cellRange = /^(((([a-zA-Z]+)|([$][a-zA-Z]+))(([0-9]+)|([$][0-9]+)))|((([a-zA-Z]+)|([$][a-zA-Z]+))))$/g; //增加正则判断单元格为字母+数字或字母的格式：如 A1:B3，A:A

    if (rangetxt.indexOf(":") == -1) {
      var row = parseInt(rangetxt.replace(/[^0-9]/g, "")) - 1;
      var col = ABCatNum(rangetxt.replace(/[^A-Za-z]/g, ""));

      if (!isNaN(row) && !isNaN(col) && rangetxt.toString().match(reg_cell)) {
        return true;
      } else if (!isNaN(row)) {
        return false;
      } else if (!isNaN(col)) {
        return false;
      } else {
        return false;
      }
    } else {
      reg_cellRange = /^(((([a-zA-Z]+)|([$][a-zA-Z]+))(([0-9]+)|([$][0-9]+)))|((([a-zA-Z]+)|([$][a-zA-Z]+)))|((([0-9]+)|([$][0-9]+s))))$/g;
      rangetxt = rangetxt.split(":");
      var row = [],
          col = [];
      row[0] = parseInt(rangetxt[0].replace(/[^0-9]/g, "")) - 1;
      row[1] = parseInt(rangetxt[1].replace(/[^0-9]/g, "")) - 1;

      if (row[0] > row[1]) {
        return false;
      }

      col[0] = ABCatNum(rangetxt[0].replace(/[^A-Za-z]/g, ""));
      col[1] = ABCatNum(rangetxt[1].replace(/[^A-Za-z]/g, ""));

      if (col[0] > col[1]) {
        return false;
      }

      if (rangetxt[0].toString().match(reg_cellRange) && rangetxt[1].toString().match(reg_cellRange)) {
        return true;
      } else {
        return false;
      }
    }
  };

  fromulaRef.isfreezonFuc = function (txt) {
    var row = txt.replace(/[^0-9]/g, "");
    var col = txt.replace(/[^A-Za-z]/g, "");
    var row$ = txt.substr(txt.indexOf(row) - 1, 1);
    var col$ = txt.substr(txt.indexOf(col) - 1, 1);
    var ret = [false, false];

    if (row$ == "$") {
      ret[0] = true;
    }

    if (col$ == "$") {
      ret[1] = true;
    }

    return ret;
  };

  fromulaRef.operator = '==|!=|<>|<=|>=|=|+|-|>|<|/|*|%|&|^';
  fromulaRef.error = {
    v: "#VALUE!",
    n: "#NAME?",
    na: "#N/A",
    r: "#REF!",
    d: "#DIV/0!",
    nm: "#NUM!",
    nl: "#NULL!",
    sp: "#SPILL!" //数组范围有其它值

  };
  fromulaRef.operatorjson = null;
  return fromulaRef;
}();

exports.fromulaRef = fromulaRef;

function isChinese(temp) {
  var re = /[^\u4e00-\u9fa5]/;
  var reg = /[\u3002|\uff1f|\uff01|\uff0c|\u3001|\uff1b|\uff1a|\u201c|\u201d|\u2018|\u2019|\uff08|\uff09|\u300a|\u300b|\u3008|\u3009|\u3010|\u3011|\u300e|\u300f|\u300c|\u300d|\ufe43|\ufe44|\u3014|\u3015|\u2026|\u2014|\uff5e|\ufe4f|\uffe5]/;
  if (reg.test(temp)) return true;
  if (re.test(temp)) return false;
  return true;
}

exports.isChinese = isChinese;

function isJapanese(temp) {
  var re = /[^\u0800-\u4e00]/;
  if (re.test(temp)) return false;
  return true;
}

exports.isJapanese = isJapanese;

function isKoera(chr) {
  if (chr > 0x3130 && chr < 0x318F || chr >= 0xAC00 && chr <= 0xD7A3) {
    return true;
  }

  return false;
}

exports.isKoera = isKoera;

function isContainMultiType(str) {
  var isUnicode = false;

  if (escape(str).indexOf("%u") > -1) {
    isUnicode = true;
  }

  var isNot = false;
  var reg = /[0-9a-z]/gi;

  if (reg.test(str)) {
    isNot = true;
  }

  var reEnSign = /[\x00-\xff]+/g;

  if (reEnSign.test(str)) {
    isNot = true;
  }

  if (isUnicode && isNot) {
    return true;
  }

  return false;
}

exports.isContainMultiType = isContainMultiType;

function getBinaryContent(path, options) {
  var promise, resolve, reject;
  var callback;

  if (!options) {
    options = {};
  } // taken from jQuery


  var createStandardXHR = function createStandardXHR() {
    try {
      return new window.XMLHttpRequest();
    } catch (e) {}
  };

  var createActiveXHR = function createActiveXHR() {
    try {
      return new window.ActiveXObject("Microsoft.XMLHTTP");
    } catch (e) {}
  }; // Create the request object


  var createXHR = typeof window !== "undefined" && window.ActiveXObject ?
  /* Microsoft failed to properly
  * implement the XMLHttpRequest in IE7 (can't request local files),
  * so we use the ActiveXObject when it is available
  * Additionally XMLHttpRequest can be disabled in IE7/IE8 so
  * we need a fallback.
  */
  function () {
    return createStandardXHR() || createActiveXHR();
  } : // For all other browsers, use the standard XMLHttpRequest object
  createStandardXHR; // backward compatible callback

  if (typeof options === "function") {
    callback = options;
    options = {};
  } else if (typeof options.callback === 'function') {
    // callback inside options object
    callback = options.callback;
  }

  resolve = function resolve(data) {
    callback(null, data);
  };

  reject = function reject(err) {
    callback(err, null);
  };

  try {
    var xhr = createXHR();
    xhr.open('GET', path, true); // recent browsers

    if ("responseType" in xhr) {
      xhr.responseType = "arraybuffer";
    } // older browser


    if (xhr.overrideMimeType) {
      xhr.overrideMimeType("text/plain; charset=x-user-defined");
    }

    xhr.onreadystatechange = function (event) {
      // use `xhr` and not `this`... thanks IE
      if (xhr.readyState === 4) {
        if (xhr.status === 200 || xhr.status === 0) {
          try {
            resolve(function (xhr) {
              // for xhr.responseText, the 0xFF mask is applied by JSZip
              return xhr.response || xhr.responseText;
            }(xhr));
          } catch (err) {
            reject(new Error(err));
          }
        } else {
          reject(new Error("Ajax error for " + path + " : " + this.status + " " + this.statusText));
        }
      }
    };

    if (options.progress) {
      xhr.onprogress = function (e) {
        options.progress({
          path: path,
          originalEvent: e,
          percent: e.loaded / e.total * 100,
          loaded: e.loaded,
          total: e.total
        });
      };
    }

    xhr.send();
  } catch (e) {
    reject(new Error(e), null);
  } // returns a promise or undefined depending on whether a callback was
  // provided


  return promise;
}

exports.getBinaryContent = getBinaryContent;
/**
 * multi sequence conversion
 * example:
 *  1、E14 -> 13_4
 *  2、E14 J14 O14 T14 Y14 AD14 AI14 AN14 AS14 AX14 ->
 *     ['13_4', '13_9','13_14', '13_19', '13_24', '13_3', '13_8',  '13_13', '13_18', '13_23']
 *  3、E46:E47 -> ['45_4',  '46_4']
 *
 * @param {string} sqref - before sequence
 * @returns {string[]}
 */

function getMultiSequenceToNum(sqref) {
  if (!sqref || (sqref === null || sqref === void 0 ? void 0 : sqref.length) <= 0) return [];
  sqref = sqref.toUpperCase();
  var sqrefRawArr = sqref.split(" ");
  var sqrefArr = sqrefRawArr.filter(function (e) {
    return e && e.trim();
  });
  var sqrefLastArr = getSqrefRawArrFormat(sqrefArr);
  var resArr = [];

  for (var i = 0; i < sqrefLastArr.length; i++) {
    var _res = getSingleSequenceToNum(sqrefLastArr[i]);

    if (_res) resArr.push(_res);
  }

  return resArr;
}

exports.getMultiSequenceToNum = getMultiSequenceToNum;
/**
 * get region sequence
 * example:
 *  1、[A1:C2'] -> ['A1', 'A2', 'B1', 'B2', 'C1', 'C2']
 *
 * @param {string[]} arr - formats arr
 * @returns {string[]} - after arr
 */

function getRegionSequence(arr) {
  var _a, _b;

  var formatArr = [];
  var regEn = new RegExp(/[A-Z]+|[0-9]+/g);
  var startArr = (_a = arr[0]) === null || _a === void 0 ? void 0 : _a.match(regEn);
  var lastArr = (_b = arr[1]) === null || _b === void 0 ? void 0 : _b.match(regEn);
  var columnMax = Math.max.apply(Math, [ABCatNum(startArr[0]), ABCatNum(lastArr[0])]);
  var columnMin = Math.min.apply(Math, [ABCatNum(startArr[0]), ABCatNum(lastArr[0])]);
  var rowMax = Math.max.apply(Math, [parseInt(startArr[1]), parseInt(lastArr[1])]);
  var rowMin = Math.min.apply(Math, [parseInt(startArr[1]), parseInt(lastArr[1])]);

  for (var i = columnMin; i <= columnMax; i++) {
    for (var j = rowMin; j <= rowMax; j++) {
      formatArr.push("" + chatatABC(i) + j);
    }
  }

  return formatArr;
}

exports.getRegionSequence = getRegionSequence;
/**
 * unified processing of conversion formats
 * example:
 *  1、['E38', 'A1:C2'] -> ['E38', 'A1', 'A2', 'B1', 'B2', 'C1', 'C2']
 *
 * @param {string[]} arr - formats arr
 * @returns {string[]} - after arr
 */

function getSqrefRawArrFormat(arr) {
  arr === null || arr === void 0 ? void 0 : arr.map(function (el) {
    if (el.includes(":")) {
      var tempArr = el.split(":");

      if ((tempArr === null || tempArr === void 0 ? void 0 : tempArr.length) === 2) {
        arr = arr.concat(getRegionSequence(tempArr));
        arr.splice(arr.indexOf(el), 1);
      }
    }
  });
  var resultArr = arr.filter(function (value, index, array) {
    return array.indexOf(value) === index;
  });
  return resultArr;
}

exports.getSqrefRawArrFormat = getSqrefRawArrFormat;
/**
 * single sequence to number
 * example:
 *  1、A1 -> 0_0
 *  2、ES14 -> 13_4
 *
 * @param {string} sqref - before sequence
 * @returns {string} - after sequence
 */

function getSingleSequenceToNum(sqref) {
  var sqrefArray = sqref.match(/[A-Z]+|[0-9]+/g);
  var sqrefLen = sqrefArray.length;
  var regEn = new RegExp("^[A-Z]+$");
  var ret = "";

  for (var i = sqrefLen - 1; i >= 0; i--) {
    var cur = sqrefArray[i];

    if (regEn.test(cur)) {
      ret += ABCatNum(cur) + "_";
    } else {
      ret += parseInt(cur) - 1 + "_";
    }
  }

  return ret.substring(0, ret.length - 1);
}

exports.getSingleSequenceToNum = getSingleSequenceToNum;
/**
 * R1C1 to Sequence
 * example: sheet2!R1C1 => sheet!A1
 *
 * @param {string} value - R1C1 value
 * @returns
 */

function getTransR1C1ToSequence(value) {
  if (!value && (value === null || value === void 0 ? void 0 : value.length) <= 0) return "";
  var len = value.length;
  var index = value.lastIndexOf("!");
  var valueArr = [value.slice(0, index), value.slice(index + 1, len)];
  var repStr = valueArr[1] || "";
  var indexR = repStr.indexOf("R");
  var indexC = repStr.indexOf("C");
  var row = Number(repStr.slice(indexR + 1, indexC));
  var column = chatatABC(Number(repStr.slice(indexC + 1, repStr === null || repStr === void 0 ? void 0 : repStr.length)) - 1);
  return valueArr[0] + "!" + column + row;
}

exports.getTransR1C1ToSequence = getTransR1C1ToSequence;
/**
 * strip x14 format data
 *
 * @param {string} value
 * @returns {Object} - { formula, sqref }
 */

function getPeelOffX14(value) {
  var _a;

  if (!value || (value === null || value === void 0 ? void 0 : value.length) <= 0) return {}; // formula

  var formulaReg = new RegExp("</x14:formula[^]>", "g");
  var lastIndex = (_a = value.match(formulaReg)) === null || _a === void 0 ? void 0 : _a.length;
  var lastValue = "</x14:formula" + lastIndex + ">";
  var lastValueEnd = value.indexOf(lastValue);
  var formulaValue = value.substring(0, lastValueEnd + lastValue.length);
  formulaValue = formulaValue.replace(/<xm:f>/g, "").replace(/<\/xm:f>/g, "").replace(/x14:/g, "").replace(/\/x14:/g, "");
  var formula = formulaValue; // sqref

  var xmSqrefLen = "<xm:sqref>".length;
  var sqrefStart = value.indexOf("<xm:sqref>");
  var sqrefEnd = value.indexOf("</xm:sqref>");
  var sqref = value.substring(sqrefStart + xmSqrefLen, sqrefEnd);
  return {
    formula: formula,
    sqref: sqref
  };
}

exports.getPeelOffX14 = getPeelOffX14;
/**
 * get the value in the formula
 *
 * @param {string} value - extracted value
 * @returns {string[]}
 */

function getMultiFormulaValue(value) {
  var _a, _b;

  if (!value || (value === null || value === void 0 ? void 0 : value.length) <= 0) return [];
  var lenReg = new RegExp("formula", "g");
  var len = (((_a = value.match(lenReg)) === null || _a === void 0 ? void 0 : _a.length) || 0) / 2;
  if (len === 0) return [];
  var retArr = [];

  for (var i = 1; i <= len; i++) {
    var startLen = (_b = "<formula" + i + ">") === null || _b === void 0 ? void 0 : _b.length;
    var start = value.indexOf("<formula" + i + ">");
    var end = value.indexOf("</formula" + i + ">");

    var _value = value.substring(start + startLen, end);

    retArr.push(escapeCharacter(_value.replace(/&quot;/g, "")));
  }

  return retArr;
}

exports.getMultiFormulaValue = getMultiFormulaValue;

},{"./constant":16}],19:[function(require,module,exports){
"use strict";

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.LuckyExcel = void 0;

var LuckyFile_1 = require("./ToLuckySheet/LuckyFile"); // import {SecurityDoor,Car} from './content';


var HandleZip_1 = require("./HandleZip"); // //demo
// function demoHandler(){
//     let upload = document.getElementById("Luckyexcel-demo-file");
//     let selectADemo = document.getElementById("Luckyexcel-select-demo");
//     let downlodDemo = document.getElementById("Luckyexcel-downlod-file");
//     let mask = document.getElementById("lucky-mask-demo");
//     if(upload){
//         window.onload = () => {
//             upload.addEventListener("change", function(evt){
//                 var files:FileList = (evt.target as any).files;
//                 if(files==null || files.length==0){
//                     alert("No files wait for import");
//                     return;
//                 }
//                 let name = files[0].name;
//                 let suffixArr = name.split("."), suffix = suffixArr[suffixArr.length-1];
//                 if(suffix!="xlsx"){
//                     alert("Currently only supports the import of xlsx files");
//                     return;
//                 }
//                 LuckyExcel.transformExcelToLucky(files[0], function(exportJson:any, luckysheetfile:string){
//                     if(exportJson.sheets==null || exportJson.sheets.length==0){
//                         alert("Failed to read the content of the excel file, currently does not support xls files!");
//                         return;
//                     }
//                     console.log(exportJson, luckysheetfile);
//                     window.luckysheet.destroy();
//                     window.luckysheet.create({
//                         container: 'luckysheet', //luckysheet is the container id
//                         showinfobar:false,
//                         data:exportJson.sheets,
//                         title:exportJson.info.name,
//                         userInfo:exportJson.info.name.creator
//                     });
//                 });
//             });
//             selectADemo.addEventListener("change", function(evt){
//                 var obj:any = selectADemo;
//                 var index = obj.selectedIndex;
//                 var value = obj.options[index].value;
//                 var name = obj.options[index].innerHTML;
//                 if(value==""){
//                     return;
//                 }
//                 mask.style.display = "flex";
//                 LuckyExcel.transformExcelToLuckyByUrl(value, name, function(exportJson:any, luckysheetfile:string){
//                     if(exportJson.sheets==null || exportJson.sheets.length==0){
//                         alert("Failed to read the content of the excel file, currently does not support xls files!");
//                         return;
//                     }
//                     console.log(exportJson, luckysheetfile);
//                     mask.style.display = "none";
//                     window.luckysheet.destroy();
//                     window.luckysheet.create({
//                         container: 'luckysheet', //luckysheet is the container id
//                         showinfobar:false,
//                         data:exportJson.sheets,
//                         title:exportJson.info.name,
//                         userInfo:exportJson.info.name.creator
//                     });
//                 });
//             });
//             downlodDemo.addEventListener("click", function(evt){
//                 var obj:any = selectADemo;
//                 var index = obj.selectedIndex;
//                 var value = obj.options[index].value;
//                 if(value.length==0){
//                     alert("Please select a demo file");
//                     return;
//                 }
//                 var elemIF:any = document.getElementById("Lucky-download-frame");
//                 if(elemIF==null){
//                     elemIF = document.createElement("iframe");
//                     elemIF.style.display = "none";
//                     elemIF.id = "Lucky-download-frame";
//                     document.body.appendChild(elemIF);
//                 }
//                 elemIF.src = value;
//                 // elemIF.parentNode.removeChild(elemIF);
//             });
//         }
//     }
// }
// demoHandler();
// api


var LuckyExcel =
/** @class */
function () {
  function LuckyExcel() {}

  LuckyExcel.transformExcelToLucky = function (excelFile, callBack) {
    var handleZip = new HandleZip_1.HandleZip(excelFile);
    handleZip.unzipFile(function (files) {
      var luckyFile = new LuckyFile_1.LuckyFile(files, excelFile.name);
      var luckysheetfile = luckyFile.Parse();
      var exportJson = JSON.parse(luckysheetfile);

      if (callBack != undefined) {
        callBack(exportJson, luckysheetfile);
      }
    }, function (err) {
      console.error(err);
    });
  };

  LuckyExcel.transformExcelToLuckyByUrl = function (url, name, callBack) {
    var handleZip = new HandleZip_1.HandleZip();
    handleZip.unzipFileByUrl(url, function (files) {
      var luckyFile = new LuckyFile_1.LuckyFile(files, name);
      var luckysheetfile = luckyFile.Parse();
      var exportJson = JSON.parse(luckysheetfile);

      if (callBack != undefined) {
        callBack(exportJson, luckysheetfile);
      }
    }, function (err) {
      console.error(err);
    });
  };

  LuckyExcel.transformLuckyToExcel = function (LuckyFile, callBack) {};

  return LuckyExcel;
}();

exports.LuckyExcel = LuckyExcel;

},{"./HandleZip":9,"./ToLuckySheet/LuckyFile":12}],20:[function(require,module,exports){
"use strict";

var main_1 = require("./main");

module.exports = main_1.LuckyExcel;

},{"./main":19}]},{},[20])(20)
});
