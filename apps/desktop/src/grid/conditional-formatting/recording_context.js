export class RecordingContext2D {
  constructor() {
    this.commands = [];
    this._fillStyle = undefined;
  }

  set fillStyle(v) {
    this._fillStyle = v;
    this.commands.push(["fillStyle", v]);
  }
  get fillStyle() {
    return this._fillStyle;
  }

  fillRect(x, y, w, h) {
    this.commands.push(["fillRect", x, y, w, h]);
  }

  beginPath() {
    this.commands.push(["beginPath"]);
  }
  moveTo(x, y) {
    this.commands.push(["moveTo", x, y]);
  }
  lineTo(x, y) {
    this.commands.push(["lineTo", x, y]);
  }
  closePath() {
    this.commands.push(["closePath"]);
  }
  fill() {
    this.commands.push(["fill"]);
  }
}
