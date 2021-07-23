let PptxGenJS = require("pptxgenjs");
var dateFormat = require("dateformat");

class ampcxstory {
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
    this.filename = project.ProjectCustomerName.replace(" ", "-") + "-AMP" + "-" + today + ".pptx";
  
  }

  generate() {
    this.pptx.layout = "LAYOUT_WIDE";
    this.addMaster();
    var slide = this.pptx.addSlide(this.masterPlaceholder);
    this.addTitleBar(slide, this.project);
    this.addNotesHeaders(slide, this.project);
    this.addNotesBodies(slide, this.project);
    this.addFooters(slide, this.project);
    slide.addNotes(
      'Customer Overview and Objective\nThis section needs to briefly outline:\nWho the customer is and what their business/mission is\nKey and/or relevant information on relationship/opps with MSFT prior to AMP\nthe customer’s objective, the "WHY" .. Why were they pursuing this change, and what was the ideal outcome?\nTell a brief story around the Why and What – FTA PM should be able to pull a lot of this content from the Project Overview in Ceres or your scoping notes prior to project creation\n\nAzure Solution Overview\nFTA PM work with FTA Engineer to populate this section if you are unsure. Provide a high-level overview of the solution we helped them implement, as well as key services and components of the FTA delivery (ie, Governance discussions, Arch Design Review, over-the-shoulder working sessions, etc…)\n\nBusiness Results and Impact\nThis is one of the most critical pieces of the Business Impact Overview. Most audiences viewing this are going to want to know what the Customer Impact was and what it has enabled the customer to do next. You can also touch on the impact to Microsoft (ACR? ACA win?, etc..) There should always be something to tell here in terms of “What’s Next” – maybe they are moving additional machines to Azure, planning modernization efforts over the next 6 months, evaluating supplemental services, etc…'
    );
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
      w: 1.1,
      h: 0.77,
      color: this.COLOR_WHITE,
      fontFace: "Segoe UI",
      fontSize: 16,
      align: "center",
      fill: { color: this.COLOR_BOX_BACKGROUND },
      line: { color: "FFFFFF", width: 1 },
    });

    slide.addText(project.ProjectCustomerName, {
      shape: this.pptx.shapes.RECTANGLE,
      autoFit: true,
      y: 0.13,
      x: 1.3,
      w: 5.4,
      h: 0.6,
      color: this.COLOR_WHITE,
      fontFace: "Segoe UI Semibold",
      fontSize: 24,
      align: "left",
      valign: "top",
    });

   slide.addImage({
      path: "img/quotes.png",
      y: 0.18,
      x: 6.7,
      w: 0.35,
      h: 0.34,
    });

    slide.addText("<Insert customer quote here>", {
      shape: this.pptx.shapes.RECTANGLE,
      y: 0.21,
      x: 7,
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
      y: 0.65,
      x: 10.85,
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
    slide.addText(" Customer Overview and Objective", {
      shape: this.pptx.shapes.RECTANGLE,
      y: 1,
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

    slide.addText(" Azure Solution Overview", {
      shape: this.pptx.shapes.RECTANGLE,
      autoFit: false,
      y: 1,
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

    slide.addText(" Business Results and Impact", {
      shape: this.pptx.shapes.RECTANGLE,
      autoFit: false,
      y: 1,
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
    var top = 1.42;

    //Customer Overview and Objective

    slide.addText(
      [
        {
          text: "Brief “About” the customer (who they are, what their business/mission is)",
          options: { italic: true, bullet: false, breakLine: true, fontSize: 11 },
        },
        {
          text: "What was the customer’s objective?",
          options: { bullet: { indent: 10 } },
        },
        {
          text: "How did this fit into their cloud strategy?",
          options: { bullet: { indent: 10 } },
        },
        {
          text: "Business drivers/justification",
          options: { bullet: { indent: 10 } },
        },
        {
          text:
            "Ideal customer outcome - How would this improve their daily/weekly/monthly business functions",
          options: { bullet: { indent: 10 } },
        },
        {
          text:
            "How large is their estate? (# of VMs/DBs/Users)",
          options: { bullet: { indent: 10 } },
        },
        {
          text:
            "What does the current estate look like (source machines?)",
          options: { bullet: { indent: 10 } },
        },
      ],
      {
        y: top,
        x: 0.12,
        w: 3.49,
        h: 3.7,
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
          text: "#VMs/DBs/Users migrated within the AMP Project",
          options: { bullet: { indent: 10 } },
        },
        {
          text: "What were the key components of the delivery?",
          options: { bullet: { indent: 10 } },
        },
        {
          text: "Working Sessions, Architecture Design Review, etc…",
          options: { bullet: { indent: 10 }, indentLevel: 2 },
        },
        {
          text: "What was implemented to meet customer’s objective?",
          options: { bullet: { indent: 10 } },
        },
        {
          text:
            "Please note specific products and services used, but refrain from simply a list",
          options: { bullet: { indent: 10 }, indentLevel: 2 },
        },
        {
          text: "Were there any technical risks or challenges that occurred and how were they resolved?",
          options: { bullet: { indent: 10 } },
        },
        {
          text: "For DC Migration: where were their source machines hosted (VMWare, Hyper-V, Physical, AWS, GCP) and what was the destination?",
          options: { bullet: { indent: 10 } },
        },
      ],
      {
        y: top,
        x: 3.71,
        h: 2.8,
        w: 4.7,
        shrinkText: true,
        fill: this.COLOR_BOX_BACKGROUND,
        color: this.COLOR_WHITE,
        fontFace: "Segoe UI",
        fontSize: 11,
        valign: "top",
      }
    );

    slide.addImage({
      path: "img/acrchart.png",
      y: 4.3,
      x: 3.71,
      h: 2.32,
      w: 4.7,
    });

    slide.addText(
      [
        {
          text: "SPOKE PM:",
          options: { bold: true, fontSize: 10, align: "center" },
        },
        {
          text:
            "Step 1: Turn on story flag in Ceres, choose success or learning.\n",
        },
        { text: 
            "Step 2: Provide the customer TPID here -> #########\n" },
        {
          text:
            "Once the story is ready for publishing the Hub PM will add the waterfall ACR graph from Customer PBI",
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
            color: "00B0F0",
          },
        },
        {
          text: "How did we directly impact the customer?\n",
          options: { bullet: { indent: 10 } },
        },
        {
          text: "Was there a reduction in time/cost/effort as a result?\n",
          options: { bullet: { indent: 10 } },
        },
        {
          text: "Was this an ACA win or other impact for Microsoft?\n",
          options: { bullet: { indent: 10 } },
        },
        {
          text: "ACR pre-FTA involvement, during, and post\n",
          options: { bullet: { indent: 10 } },
        },
      ],
      {
        autoFit: false,
        y: top,
        x: 8.51,
        w: 4.7,
        h: 5.2,
        shrinkText: true,
        fontFace: "Segoe UI",
        fontSize: 11,
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
            color: "00B0F0",
          },
        },
        {
          text: "What specific product feedback was captured?\n",
          options: { bullet: { indent: 10 } },
        },
        {
          text: "What were the deltas in FTA NPS/NSAT?\n",
          options: { bullet: { indent: 10 } },
        },
        {
          text: "What challenges were encountered and how were they overcome?",
          options: { bullet: { indent: 10 } },
        },
      ],
      {
        autoFit: false,
        y: 2.7,
        x: 8.51,
        w: 4.7,
        h: 2,
        shrinkText: true,
        fontFace: "Segoe UI",
        fontSize: 11,
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
            color: "00B0F0",
          },
        },
        {
          text: "What is next for the customer?\n",
          options: { bullet: { indent: 10 } },
        },
        {
          text: "Any optimization/tuning efforts planned?\n",
          options: { bullet: { indent: 10 } },
        },
        {
          text: "Are there any other initiatives FTA can assist with?",
          options: { bullet: { indent: 10 } },
        },
      ],
      {
        autoFit: false,
        y: 4,
        x: 8.51,
        w: 4.7,
        h: 2,
        shrinkText: true,
        fontFace: "Segoe UI",
        fontSize: 11,
        color: "FFFFFF",
        valign: "top",
        align: "left",
        fill: this.COLOR_BOX_BACKGROUND,
      }
    );

    //Collaborators
    slide.addText(
      [
        { text: "Collaborators\n", options: { bold: true, fontSize: 12, color: "00B0F0", italic: true, } },
        { text: "Partner: ", options: { bold: true } },
        { text: "Name\n", options: { bold: false } },
        { text: "Field: ", options: { bold: true } },
        { text: "Name (Role)\n", options: { bold: false } },
        { text: "EPM: (Remove if N/A) ", options: { bold: true } },
        { text: "Name (Role)\n", options: { bold: false } },
        { text: "FastTrack for Azure: ", options: { bold: true } },
        { text: "Name (Role)\n", options: { bold: false } },
        { text: "Skilling: (Remove if N/A) ", options: { bold: true } },
        { text: "Name (Role)\n", options: { bold: false } },
      ],
      {
      shape: this.pptx.shapes.RECTANGLE,
      autoFit: false,
      y: 5.2,
      x: 0.12,
      w: 3.49,
      h: 1.41,
      fill: this.COLOR_BOX_BACKGROUND,
      fontFace: "Segoe UI",
      fontSize: 9,
      color: "FFFFFF",
      align: "left",
      valign: "top",
    });


  }


    // footer

  addFooters(slide, project) {
    var footerY = 6.7;
    var footerTxtY = 6.71;
    var footerH = 0.6;
    var imageH = 0.6;
    var imageW = 0.6;

    var startDate = dateFormat(project.StartDate, "yyyy-mm-dd");
    var endDate = dateFormat(project.EndDate, "yyyy-mm-dd");

    // footer wrapper

    slide.addText(" ", {
      shape: this.pptx.shapes.RECTANGLE,
      y: footerY,
      x: 0,
      h: footerH,
      w: 13.33,
      fill: "50E6FF",
      color: "A6A6A6",
      fontFace: "Segoe UI",
      fontSize: 11,
      align: "center",
      valign: "bottom",
    });

    //Customer Location

    slide.addImage({
      path: "img/location.png",
      y: footerY,
      x: 0.45,
      h: imageH,
      w: imageW,
    });

    slide.addText(
      [
        { text: "Customer Location:\n", options: { bold: true } },
        { text: project.PhysicalLocation, options: { bold: false } },
      ],
      {
        y: footerTxtY,
        x: 1,
        h: footerH,
        w: 1.85,
        valign: "top",
        fontFace: "Segoe UI",
        fontSize: 11,
        color: "000000",
      }
    );

    //Customer Industry

    slide.addImage({
      path: "img/industry.png",
      y: footerY,
      x: 2.91,
      h: imageH,
      w: imageW,
    });

    slide.addText(
      [
        { text: "Customer Industry: \n", options: { bold: true } },
        { text: project.QualifiedIndustry, options: { bold: false } },
      ],
      {
        y: footerTxtY,
        x: 3.49,
        h: footerH,
        w: 2,
        valign: "top",
        fontFace: "Segoe UI",
        fontSize: 11,
        color: "000000",
      }
    );

        //Customer Segment

        slide.addImage({
          path: "img/solution.png",
          y: footerY,
          x: 5.55,
          h: imageH,
          w: imageW,
        });
    
        slide.addText(
          [
            { text: "Customer Segment: ", options: { bold: true } },
            { text: project.QualifiedSegment, options: { bold: false } },
              ],
          {
            y: footerTxtY,
            x: 6.06,
            h: footerH,
            w: 2,
            valign: "top",
            fontFace: "Segoe UI",
            fontSize: 11,
            color: "000000",
          }
        );
    
    //Partner

    slide.addImage({
      path: "img/partner.png",
      y: footerY,
      x: 8.1,
      h: imageH,
      w: imageW,
});

    slide.addText(
      [
        { text: "Partner: \n", options: { bold: true } },
        { text: "<type name here>", options: { bold: false } },
      ],
      {
        y: footerTxtY,
        x: 8.63,
        h: footerH,
        w: 2,
        valign: "top",
        fontFace: "Segoe UI",
        fontSize: 11,
        color: "000000",
      }
    );

    //Project Duration

    slide.addImage({
      path: "img/duration.png",
      y: footerY,
      x: 10.57,
      h: imageH,
      w: imageW,
    });

    slide.addText(
      [
        { text: "Project Duration: \n", options: { bold: true } },
        { text: startDate + " – " + endDate, options: { bold: false } },
      ],
      {
        y: footerTxtY,
        x: 11.02,
        h: footerH,
        w: 2,
        valign: "top",
        fontFace: "Segoe UI",
        fontSize: 11,
        color: "000000",
      }
    );


    // MS Confidential

        slide.addText(
          [
            {text: "MICROSOFT CONFIDENTIAL - INTERNAL ONLY",},
          ],
          {
            autoFit: false,
            y: 7.32,
            x: 0,
            w: 13.3,
            h: .22,
            fontFace: "Segoe UI",
            fontSize: 11,
            color: "A6A6A6",
            valign: "bottom",
            align: "center",
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
  ampcxstory: ampcxstory,
};
