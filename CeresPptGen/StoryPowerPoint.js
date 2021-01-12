let PptxGenJS = require("pptxgenjs");

class StoryPowerPoint {
  constructor(project) {
    this.project = project;
    this.pptx = new PptxGenJS();
    this.masterPlaceholder = "PLACEHOLDER_SLIDE";

    //Pseudo constants
    this.COLOR_WHITE = "FFFFFF";
    this.COLOR_BLACK = "000000";
    this.COLOR_RED = "fc0303";
    this.COLOR_GREEN = "00b04f";
    this.COLOR_BLUE = "4472c4";
    this.COLOR_DARK_BLUE = "44546a";
    this.COLOR_AMBER = "ffbf00";
    this.COLOR_GRAY = "808080";
    this.COLOR_LIGHT_GRAY = "f2f2f2";
    this.COLOR_NOTES_HEADER = "0078D4";
    this.COLOR_BOX_BACKGROUND = "262626";
    this.COLOR_CELL_HEADER = this.COLOR_WHITE; //'4473c4';
    this.COLOR_CELL_DARK = this.COLOR_LIGHT_GRAY; //'cfd5ea';
    this.COLOR_CELL_LIGHT = this.COLOR_WHITE; // 'e9ebf5';

    this.imageY = 6.7;
    this.imageH = 0.6;
    this.imageW = 0.6;

    var today = new Date().toISOString();
    today = today.replace(new RegExp(":", "g"), "-");
    this.filename = "FTA-CustomerStories-" + today + ".pptx";
  }

  generate() {
    this.pptx.layout = "LAYOUT_WIDE";
    this.addMaster();
    var slide = this.pptx.addSlide(this.masterPlaceholder);
    this.addTitleBar(slide, this.project);
    this.addNotesHeaders(slide, this.project);
    this.addNotesBodies(slide, this.project);
    this.addFooters(slide, this.project);
  }

  addMaster() {
    this.pptx.defineSlideMaster({
      title: this.masterPlaceholder,
      bkgd: this.COLOR_BLACK,
    });
  }

  addTitleBar(slide, project) {
    slide.addText("Logo\nHere", {
      shape: this.pptx.shapes.ROUNDED_RECTANGLE,
      y: 0.13,
      x: 0.12,
      w: 1,
      h: 0.8,
      color: this.COLOR_WHITE,
      fontFace: "Segoe UI Semibold",
      fontSize: 16,
      align: "center",
      fill: { color: this.COLOR_BOX_BACKGROUND },
      line: { color: "FFFFFF", width: 0.5 },
    });

    slide.addText(project.NominationCustomerName, {
      shape: this.pptx.shapes.RECTANGLE,
      y: 0.13,
      x: 1.2,
      w: 5.2,
      h: 0.8,
      color: this.COLOR_WHITE,
      fontFace: "Segoe UI Semibold",
      fontSize: 24,
      align: "left",
      valign: "top",
    });

    slide.addText("Business Impact Overview", {
      shape: this.pptx.shapes.RECTANGLE,
      y: 0.6,
      x: 1.2,
      w: 2.5,
      h: 0.2,
      color: "00B0F0",
      fontFace: "Segoe UI Semibold",
      fontSize: 14,
      align: "left",
      italic: true,
    });
    slide.addImage({
      path:
        "https://cdn2.iconfinder.com/data/icons/metro-ui-icon-set/512/PowerPoint_15.png",
      y: 0.13,
      x: 6.65,
      w: 0.25,
      h: 0.25,
    });

    slide.addText("<Insert customer quote here>", {
      shape: this.pptx.shapes.RECTANGLE,
      y: 0.13,
      x: 6.85,
      w: 5.5,
      h: 0.8,
      color: this.COLOR_WHITE,
      fontFace: "Segoe UI",
      fontSize: 12,
      align: "left",
      valign: "top",
      italic: true,
    });

    slide.addText("~Name/Role/Company", {
      shape: this.pptx.shapes.RECTANGLE,
      y: 0.7,
      x: 10.5,
      w: 2,
      h: 0.2,
      color: this.COLOR_WHITE,
      fontFace: "Segoe UI",
      fontSize: 12,
      align: "left",
      italic: true,
    });
  }

  addNotesHeaders(slide, project) {
    slide.addText("Customer Overview and Objective", {
      shape: this.pptx.shapes.RECTANGLE,
      y: 1.1,
      x: 0.12,
      w: 3.49,
      h: 0.4,
      color: this.COLOR_WHITE,
      fontFace: "Segoe UI Semibold",
      fontSize: 12,
      align: "left",
      fill: { color: this.COLOR_NOTES_HEADER },
      line: { color: "FFFFFF", width: 0.5 },
    });

    slide.addText("Azure Solution Overview", {
      shape: this.pptx.shapes.RECTANGLE,
      autoFit: false,
      y: 1.1,
      x: 3.71,
      w: 4.7,
      h: 0.4,
      fill: { color: this.COLOR_NOTES_HEADER },
      line: { color: "FFFFFF", width: 0.5 },
      fontFace: "Segoe UI Semibold",
      fontSize: 12,
      color: this.COLOR_WHITE,
      align: "left",
    });

    slide.addText("Business Results and Impact", {
      shape: this.pptx.shapes.RECTANGLE,
      autoFit: false,
      y: 1.1,
      x: 8.51,
      w: 4.7,
      h: 0.4,
      fill: { color: this.COLOR_NOTES_HEADER },
      line: { color: "FFFFFF", width: 0.5 },
      fontFace: "Segoe UI Semibold",
      fontSize: 12,
      color: this.COLOR_WHITE,
      align: "left",
    });
  }

  addNotesBodies(slide, project) {
    // notes body
    var top = 1.6;
    var height = 5;

    //Customer Overview and Objective

    slide.addText(
      [
        {
          text: "What was the customer’s objective?",
          options: { bullet: { indent: 15 } },
        },
        {
          text: "How did this fit in to their cloud strategy?",
          options: { bullet: { indent: 15 } },
        },
        {
          text: "Business drivers/justification",
          options: { bullet: { indent: 15 } },
        },
        {
          text:
            "Ideal customer outcome - How would this improve their daily/weekly/monthly business functions",
          options: { bullet: { indent: 15 } },
        },
        {
          text: "If AMP, how large was their estate?",
          options: { bullet: { indent: 15 } },
        },
        {
          text: "#VMs/DBs/Users migrated",
          options: { bullet: { indent: 15 }, indentLevel: 1 },
        },
      ],
      {
        y: top,
        x: 0.12,
        w: 3.49,
        h: 3.6,
        shrinkText: true,
        fill: this.COLOR_BOX_BACKGROUND,
        color: this.COLOR_WHITE,
        fontFace: "Segoe UI",
        fontSize: 11,
        valign: "top",
      }
    );

    //Azure Solution Overview
    slide.addText(
      [
        {
          text: "What were the key components of the delivery?",
          options: { bullet: { indent: 15 } },
        },
        {
          text: "Working Sessions, Architecture Design Review, etc…",
          options: { bullet: { indent: 15 }, indentLevel: 1 },
        },
        {
          text: "What was implemented to meet customer’s objective?",
          options: { bullet: { indent: 15 } },
        },
        {
          text:
            "Please note specific products and services used, but refrain from simply a list",
          options: { bullet: { indent: 15 }, indentLevel: 1 },
        },
      ],
      {
        y: top,
        x: 3.71,
        w: 4.7,
        h: 2.6,
        shrinkText: true,
        fill: this.COLOR_BOX_BACKGROUND,
        color: this.COLOR_WHITE,
        fontFace: "Segoe UI",
        fontSize: 11,
        valign: "top",
      }
    );

    slide.addImage({
      path:
        "https://cdn2.iconfinder.com/data/icons/metro-ui-icon-set/512/PowerPoint_15.png",
      y: 4.28,
      x: 3.71,
      h: 2.32,
      w: 4.7,
    });

    slide.addText(
      [
        {
          text: "INSTRUCTIONS FOR ACR CHART",
          options: { bold: true, fontSize: 10, align: "center" },
        },
        {
          text:
            "Step 1: Turn on story flag in Ceres, choose success or learning.\n",
        },
        { text: "Step 2: data propagates (takes a few hours)" },
        {
          text:
            "Step 3: Hub story PM or project PM can scroll, select customer and snip the ACR graph from the ",
          options: { breakLine: false },
        },
        {
          text: "Story PBI",
          options: {
            hyperlink: {
              url:
                "https://msit.powerbi.com/groups/me/reports/2a4f3cc3-cd30-48fc-b043-24dfc294196a/ReportSection",
              tooltip: "FTA Story PBI",
            },
            breakLine: false,
          },
        },
        {
          text:
            ". If unmanaged you may show no $ and will need to gather from - ",
          options: { breakLine: false },
        },
        {
          text: "C+AI customer portal",
          options: {
            hyperlink: {
              url: "https://cecustomers.microsoftonline.com/",
              tooltip: "C+AI customer portal",
            },
            breakLine: false,
          },
        },
      ],
      {
        autoFit: false,
        y: 5.25,
        x: 3.9,
        h: 1.05,
        w: 4.3,
        shrinkText: true,
        fontFace: "Segoe UI",
        color: "FF0000",
        valign: "top",
        align: "left",
        fill: this.COLOR_BOX_BACKGROUND,
        fontSize: 9,
      }
    );

    //Business Results and Impact
    slide.addText(
      [
        {
          text: "Impact",
          options: {
            bold: true,
            fontSize: 12,
            italic: true,
            bullet: false,
            breakLine: true,
          },
        },
        {
          text: "How did we directly impact the customer?\n",
          options: { bullet: { indent: 15 } },
        },
        {
          text: "Was there a reduction in time/cost/effort as a result?\n",
          options: { bullet: { indent: 15 } },
        },
        {
          text: "Was this an ACA win for Microsoft?\n",
          options: { bullet: { indent: 15 } },
        },
        {
          text: "ACR pre-FTA involvement, during, and post",
          options: { bullet: { indent: 15 } },
        },
      ],
      {
        autoFit: false,
        y: top,
        x: 8.51,
        w: 4.7,
        h: height,
        shrinkText: true,
        fontFace: "Segoe UI",
        color: "FFFFFF",
        valign: "top",
        align: "left",
        fill: this.COLOR_BOX_BACKGROUND,
      }
    );

    slide.addText(
      [
        {
          text: "Lessons Learned",
          options: {
            bold: true,
            fontSize: 12,
            italic: true,
            bullet: false,
            breakLine: true,
          },
        },
        {
          text: "What specific product feedback was captured?\n",
          options: { bullet: { indent: 15 } },
        },
        {
          text: "What were the deltas in FTA NPS/NSAT?\n",
          options: { bullet: { indent: 15 } },
        },
        {
          text: "What challenges were encountered and how were they overcome?",
          options: { bullet: { indent: 15 } },
        },
      ],
      {
        autoFit: false,
        y: 3,
        x: 8.51,
        w: 4.7,
        h: height,
        shrinkText: true,
        fontFace: "Segoe UI",
        color: "FFFFFF",
        valign: "top",
        align: "left",
        fill: this.COLOR_BOX_BACKGROUND,
      }
    );

    slide.addText(
      [
        {
          text: "What's Next?",
          options: {
            bold: true,
            fontSize: 12,
            italic: true,
            bullet: false,
            breakLine: true,
          },
        },
        {
          text: "What is next for the customer?\n",
          options: { bullet: { indent: 15 } },
        },
        {
          text: "Any optimization/tuning efforts planned?\n",
          options: { bullet: { indent: 15 } },
        },
        {
          text: "Are there any other initiatives FTA can assist with?",
          options: { bullet: { indent: 15 } },
        },
      ],
      {
        autoFit: false,
        y: 4.4,
        x: 8.51,
        w: 4.7,
        h: height,
        shrinkText: true,
        fontFace: "Segoe UI",
        color: "FFFFFF",
        valign: "top",
        align: "left",
        fill: this.COLOR_BOX_BACKGROUND,
      }
    );

    //Collaborators
    slide.addText("Collaborators​\n", {
      shape: this.pptx.shapes.RECTANGLE,
      autoFit: false,
      y: 5.28,
      x: 0.12,
      w: 3.49,
      h: 1.32,
      fill: this.COLOR_BOX_BACKGROUND,
      color: "00B0F0",
      fontFace: "Segoe UI",
      fontSize: 12,
      align: "left",
      italic: true,
      valign: "top",
    });

    slide.addText(
      [
        { text: "FastTrack for Azure: ", options: { bold: true } },
        { text: "Name (Role), Name (Role)", options: { bold: false } },
      ],
      {
        y: 5.55,
        x: 0.12,
        w: 3.49,
        h: 0.2,
        shrinkText: true,
        fontFace: "Segoe UI",
        fontSize: 9,
        color: "FFFFFF",
      }
    );

    slide.addText(
      [
        { text: "Field: ", options: { bold: true } },
        { text: "Name (Role), Name (Role)", options: { bold: false } },
      ],
      {
        y: 5.7,
        x: 0.12,
        w: 3.49,
        h: 0.2,
        shrinkText: true,
        fontFace: "Segoe UI",
        fontSize: 9,
        color: "FFFFFF",
      }
    );

    slide.addText(
      [
        { text: "EPM: ", options: { bold: true } },
        { text: "AMP, remove if N/A", options: { bold: false } },
      ],
      {
        y: 5.85,
        x: 0.12,
        w: 3.49,
        h: 0.2,
        shrinkText: true,
        fontFace: "Segoe UI",
        fontSize: 9,
        color: "FFFFFF",
      }
    );

    slide.addText(
      [
        { text: "Skilling: ", options: { bold: true } },
        { text: "AMP, remove if N/A", options: { bold: false } },
      ],
      {
        y: 6,
        x: 0.12,
        w: 3.49,
        h: 0.2,
        shrinkText: true,
        fontFace: "Segoe UI",
        fontSize: 9,
        color: "FFFFFF",
      }
    );

    slide.addText(
      [
        { text: "Partner: ", options: { bold: true } },
        { text: "Indicate N/A if so", options: { bold: false } },
      ],
      {
        y: 6.15,
        x: 0.12,
        w: 3.49,
        h: 0.2,
        shrinkText: true,
        fontFace: "Segoe UI",
        fontSize: 9,
        color: "FFFFFF",
      }
    );
  }

  addFooters(slide, project) {
    // footer
    var footerY = 6.75;
    var footerH = 0.75;
    var imageH = 0.4;
    var imageW = 0.4;

    // footer wrapper and MS Confidential
    slide.addText("MICROSOFT CONFIDENTIAL - Internal Only", {
      shape: this.pptx.shapes.RECTANGLE,
      y: footerY,
      x: 0.12,
      h: footerH,
      w: 13.09,
      fill: this.COLOR_BOX_BACKGROUND,
      color: "A6A6A6",
      fontFace: "Segoe UI",
      fontSize: 11,
      align: "center",
      valign: "bottom",
    });

    //Customer Location
    slide.addImage({
      path:
        'img/country.jpg',
      y: 6.8,
      x: 0.19,
      h: imageH,
      w: imageW,
    });

    slide.addText(
      [
        { text: "Customer Location:\n", options: { bold: true } },
        { text: project.PhysicalLocation, options: { bold: false } },
      ],
      {
        y: footerY,
        x: 0.65,
        h: footerH,
        w: 2.2,
        valign: "top",
        fontFace: "Segoe UI",
        fontSize: 11,
        color: "00B0F0",
      }
    );

    slide.addImage({
      path:
      'img/duration.jpg',
      y: 6.8,
      x: 3,
      h: imageH,
      w: imageW,
    });

    slide.addText(
      [
        { text: "Customer Solution: ", options: { bold: true } },
        {
          text: "<Insert Ceres Customer Solution>\n",
          options: { bold: false },
        },
        { text: "Project Duration: ", options: { bold: true } },
        {
          text: project.StartDate + " – " + project.EndDate,
          options: { bold: false },
        },
      ],
      {
        y: footerY,
        x: 3.46,
        h: footerH,
        w: 4.5,
        valign: "top",
        fontFace: "Segoe UI",
        fontSize: 11,
        color: "00B0F0",
      }
    );

    slide.addImage({
      path:
      'img/industry.jpg',
      y: 6.8,
      x: 8.04,
      h: imageH,
      w: imageW,
    });

    slide.addText(
      [
        { text: "Customer Industry:\n", options: { bold: true } },
        { text: project.QualifiedIndustry, options: { bold: false } },
      ],
      {
        y: footerY,
        x: 8.5,
        h: footerH,
        w: 2.09,
        valign: "top",
        fontFace: "Segoe UI",
        fontSize: 11,
        color: "00B0F0",
      }
    );

    slide.addImage({
      path:
      'img/segment.jpg',
      y: 6.8,
      x: 10.63,
      h: imageH,
      w: imageW,
    });

    slide.addText(
      [
        { text: "Customer Segment:\n", options: { bold: true } },
        { text: project.QualifiedSegment, options: { bold: false } },
      ],
      {
        y: footerY,
        x: 11.09,
        h: footerH,
        w: 1.79,
        valign: "top",
        fontFace: "Segoe UI",
        fontSize: 11,
        color: "00B0F0",
      }
    );
  }

  dateFromUtc(utc) {
    //Preconditions
    if (
      !utc ||
      utc.length < 13 ||
      !Number.isInteger(parseInt(utc.substr(0, 4)))
    ) {
      //console.log("DateUtil.UTCtoMedium: utc parameter invalid: " + utc);
      return "";
    }
    //avoid "0001-01-01T08:00:00+00:00"
    if (utc.substr(0, 4) < "2000") return "";

    var year = utc.substr(0, 4);
    var month = parseInt(utc.substr(5, 2)) - 1; //Months are 0-11
    var day = utc.substr(8, 2);
    var hours = parseInt(utc.substr(11, 2));
    if (hours > 11) day++; //goes to next day, if it's more than days in the month, date object adjusts accordingly
    var dateObj = new Date(year, month, day);
    return dateObj;
  }

  daysBetween(date1, date2) {
    //Get 1 day in milliseconds
    var one_day = 1000 * 60 * 60 * 24;
    // Convert both dates to milliseconds
    var date1_ms = date1.getTime();
    var date2_ms = date2.getTime();

    // Calculate the difference in milliseconds
    var difference_ms = date2_ms - date1_ms;

    // Convert back to days and return
    return Math.round(difference_ms / one_day);
  }
}

module.exports = {
  StoryPowerPoint: StoryPowerPoint,
};
