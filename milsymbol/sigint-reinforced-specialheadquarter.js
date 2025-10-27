// sigint and reinforced addons to base geometries makes
// the displayed properties look better.

// Possible values for the options property reinforced is
// '(+)','(-)' or '(±)'

// Possible values for the options property signature is
// '' or '!'

import pathsReinforced from "./paths-reinforced";

export default function sigintReinforced(ms) {
  const drawArray1 = [];
  const drawArray2 = [];
  const bbox_org = ms.BBox(this.metadata.baseGeometry.bbox);
  const this_bbox = ms.BBox(this.bbox);
  const frameColor = this.colors.frameColor[this.metadata.affiliation];
  const fontSize = this.style.infoSize;
  const spaceTextIcon = 20;
  const stack = this.options.stack ? this.options.stack * 15 : 0;

  //If we don't have a geometry we shouldn't add anything.
  if (this.metadata.baseGeometry.g && frameColor) {
    if (
      this.metadata.affiliation == "Unknown" ||
      (this.metadata.affiliation == "Hostile" &&
        this.metadata.dimension != "Subsurface")
    ) {
    }

    if (this.options.reinforced != undefined && this.options.reinforced != "") {
      let draw = {};
      switch (this.options.reinforced) {
        case "(+)":
          draw = Object.assign({}, pathsReinforced.plus);
          draw.d =
            `M${bbox_org.x2 + spaceTextIcon + stack},${100 - 1.5 * fontSize} ` +
            draw.d;
          draw.fill = frameColor;
          drawArray2.push(draw);
          break;
        case "(-)":
          draw = Object.assign({}, pathsReinforced.minus);
          draw.d =
            `M${bbox_org.x2 + spaceTextIcon + stack},${100 - 1.5 * fontSize} ` +
            draw.d;
          draw.fill = frameColor;
          drawArray2.push(draw);
          break;
        case "(±)":
          draw = Object.assign({}, pathsReinforced.plus_minus);
          draw.d =
            `M${bbox_org.x2 + spaceTextIcon + stack},${100 - 1.5 * fontSize} ` +
            draw.d;
          draw.fill = frameColor;
          drawArray2.push(draw);
          break;
        default:
          drawArray2.push({
            type: "text",
            text: this.options.reinforced,
            x: bbox_org.x2 + spaceTextIcon,
            y: 100 - 1.5 * fontSize,
            fill: frameColor,
            fontfamily: this.style.fontfamily,
            fontsize: 50,
            //fontweight: "bold",
            textanchor: "start",
          });
      }
      this_bbox.merge({
        x2: bbox_org.x2 + spaceTextIcon + 37 + stack,
        y1: 100 - 2.5 * fontSize,
      });
    }

    if (this.options.signature == "!") {
      drawArray2.push({
        type: "text",
        text: "!",
        x: bbox_org.x2 + spaceTextIcon + stack /*+ flag*/,
        y: 100 + 2.5 * fontSize,
        fill: frameColor,
        fontfamily: this.style.fontfamily,
        fontsize: fontSize,
        fontweight: "bold",
        textanchor: "start",
      });
      this_bbox.merge({
        x2: bbox_org.x2 + spaceTextIcon + 22 + stack,
        y1: 170 - 25,
        y2: 100 + 2.5 * fontSize,
      });
    }
    if (
      this.options.specialheadquarter != undefined &&
      this.options.specialheadquarter != ""
    ) {
      let size = 45;
      const fontFamily = this.style.fontfamily;
      const fontColor =
        (typeof this.style.infoColor === "object"
          ? this.style.infoColor[this.metadata.affiliation]
          : this.style.infoColor) ||
        this.colors.iconColor[this.metadata.affiliation] ||
        this.colors.iconColor["Friend"];
      const y = 103;
      const str = this.options.specialheadquarter;
      if (str.length == 1) {
        size = 45;
      }
      if (str.length == 3) {
        size = 35;
      }
      if (str.length == 4) {
        size = 32;
      }
      if (str.length == 5) {
        size = 29;
      }
      if (str.length == 6) {
        size = 26;
      }
      if (str.length == 7) {
        size = 25;
      }
      if (str.length >= 8) {
        size = 24;
      }

      drawArray2.push({
        type: "text",
        text: this.options.specialheadquarter,
        x: 100,
        y: y,
        textanchor: "middle",
        alignmentBaseline: "middle",
        fontsize: size,
        fontfamily: fontFamily,
        fill: fontColor,
        stroke: false,
        fontweight: "bold",
      });
    }

    //outline
    if (this.style.outlineWidth > 0)
      drawArray1.push(
        ms.outline(
          drawArray2,
          this.style.outlineWidth,
          this.style.strokeWidth,
          typeof this.style.outlineColor === "object"
            ? this.style.outlineColor[this.metadata.affiliation]
            : this.style.outlineColor,
        ),
      );
  }
  return { pre: drawArray1, post: drawArray2, bbox: this_bbox };
}
