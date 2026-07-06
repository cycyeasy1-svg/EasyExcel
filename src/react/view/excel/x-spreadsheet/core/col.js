import helper from './helper';

function normalizeScale(scale) {
  const n = Number(scale);
  return Number.isFinite(n) && n > 0 ? n : 1;
}

class Cols {
  constructor({
    len, width, indexWidth, minWidth,
  }) {
    this._ = {};
    this.len = len;
    this.scale = 1;
    this._width = width;
    this._indexWidth = indexWidth;
    this._minWidth = minWidth;
  }

  get width() {
    return this._width * this.scale;
  }

  set width(value) {
    this._width = value / this.scale;
  }

  get indexWidth() {
    return this._indexWidth * this.scale;
  }

  set indexWidth(value) {
    this._indexWidth = value / this.scale;
  }

  get minWidth() {
    return this._minWidth * this.scale;
  }

  set minWidth(value) {
    this._minWidth = value / this.scale;
  }

  setScale(scale) {
    this.scale = normalizeScale(scale);
  }

  setData(d) {
    if (d.len) {
      this.len = d.len;
      delete d.len;
    }
    this._ = d;
  }

  getData() {
    const { len } = this;
    return Object.assign({ len }, this._);
  }

  getWidth(i) {
    if (this.isHide(i)) return 0;
    const col = this._[i];
    if (col && col.width) {
      return col.width * this.scale;
    }
    return this.width;
  }

  getOrNew(ci) {
    this._[ci] = this._[ci] || {};
    return this._[ci];
  }

  setWidth(ci, width) {
    const col = this.getOrNew(ci);
    col.width = width / this.scale;
  }

  unhide(idx) {
    let index = idx;
    while (index > 0) {
      index -= 1;
      if (this.isHide(index)) {
        this.setHide(index, false);
      } else break;
    }
  }

  isHide(ci) {
    const col = this._[ci];
    return col && col.hide;
  }

  setHide(ci, v) {
    const col = this.getOrNew(ci);
    if (v === true) col.hide = true;
    else delete col.hide;
  }

  setStyle(ci, style) {
    const col = this.getOrNew(ci);
    col.style = style;
  }

  sumWidth(min, max) {
    return helper.rangeSum(min, max, i => this.getWidth(i));
  }

  totalWidth() {
    return this.sumWidth(0, this.len);
  }
}

export default {};
export {
  Cols,
};
