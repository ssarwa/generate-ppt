let PptxGenJS = require("pptxgenjs");
var dateFormat = require("dateformat");

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
    this.filename = project.ProjectCustomerName.replace(" ", "-") + "-ArchOverview" + "-" + today + ".pptx";
  
  }

  generate() {
    this.pptx.layout = "LAYOUT_WIDE";
    this.addMaster();
    var slide = this.pptx.addSlide(this.masterPlaceholder);
    this.addTitleBar(slide, this.project);
    this.addNotesBodies(slide, this.project);
    this.addFooters(slide, this.project);
    slide.addNotes(
      '**This slide can be used in addition to Technical Overview if a larger view or the technical diagram is required, along with space for solution commentary.\nArchitecture Overview is optional but may be beneficial for complex architectures.\n\nWhen adding a Architecture Overview, please consider:\n - Readability\n - Ability for audience to understand (ie, does it illustrate services involved as opposed to just network plane?)\n\nTechnical Azure Solution\nOverall technical goal of the Project: What are the different technical components to be implemented, and why were they needed?\nGive an overview of the Technical Solution, touching on any specific customer requirements, as well as any important considerations that factored into the decision-making process.\nOutline services used; note preview services, GA offerings, or any concessions made while waiting for something to hit GA status.\nIt’s also important to discuss the Desired end-state and what the solution was enabling the customer to achieve.'
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

    slide.addText("Architecture Overview", {
      shape: this.pptx.shapes.RECTANGLE,
      y: 0.55,
      x: 1.3,
      w: 2.5,
      h: 0.2,
      color: "00B0F0",
      fontFace: "Segoe UI Semibold",
      fontSize: 14,
      align: "left",
      italic: true,
    });
  }

  addNotesBodies(slide, project) {
    // notes body
    var top = 1.28;

    //Technical Solution

    slide.addText(
      [
        {
          text: "Technical Solution",
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
          text: "Architecture Description",
          options: { bullet: { indent: 10 } },
        },
        {
          text: "Azure Services Used",
          options: { bullet: { indent: 10 } },
        },
        {
          text: "Design Considerations",
          options: { bullet: { indent: 10 } },
        },
      ],
      {
        y: top,
        x: 0.11,
        h: 5.35,
        w: 3.08,
        shrinkText: true,
        fill: this.COLOR_BOX_BACKGROUND,
        color: this.COLOR_WHITE,
        fontFace: "Segoe UI",
        fontSize: 11,
        valign: "top",
      }
    );

    //Architecture Image

    slide.addImage({
      path: "img/archdiag.png",
      y: top,
      x: 3.32,
      h: 5.35,
      w: 9.9,
    });

  }

    // footer

    addFooters(slide, project) {
      var footerY = 6.7;
      var footerTxtY = 6.72;
      var footerH = 0.6;
      var imageH = 0.5;
      var imageW = 0.5;
  
      var startDate = dateFormat(project.StartDate, "yyyy-mm-dd");
      var endDate = dateFormat(project.EndDate, "yyyy-mm-dd");
  
      // footer wrapper
  
      slide.addText(" ", {
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
        path: "img/location.png",
        y: footerY,
        x: 0.16,
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
          x: 0.61,
          h: footerH,
          w: 1.95,
          valign: "top",
          fontFace: "Segoe UI",
          fontSize: 11,
          color: "00B0F0",
        }
      );
  
      //Customer Industry & Segment
  
      slide.addImage({
        path: "img/industry.png",
        y: footerY,
        x: 2.5,
        h: imageH,
        w: imageW,
      });
  
      slide.addText(
        [
          { text: "Customer Industry: ", options: { bold: true } },
          { text: project.QualifiedIndustry + "\n", options: { bold: false } },
          { text: "Customer Segment: ", options: { bold: true } },
          { text: project.QualifiedSegment, options: { bold: false } },
        ],
        {
          y: footerTxtY,
          x: 2.95,
          h: footerH,
          w: 3.6,
          valign: "top",
          fontFace: "Segoe UI",
          fontSize: 11,
          color: "00B0F0",
        }
      );
  
      //Project Duration
  
      slide.addImage({
        path: "img/duration.png",
        y: footerY,
        x: 6.54,
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
          x: 6.99,
          h: footerH,
          w: 1.95,
          valign: "top",
          fontFace: "Segoe UI",
          fontSize: 11,
          color: "00B0F0",
        }
      );
  
      //Customer Solution & Workload Category
  
      slide.addImage({
        path: "img/solution.png",
        y: footerY,
        x: 8.88,
        h: imageH,
        w: imageW,
      });
  
      slide.addText(
        [
          { text: "Customer Solution: ", options: { bold: true } },
          { text: project.CustomerSolution + "\n", options: { bold: false } },
          { text: "Workload Category: ", options: { bold: true } },
          { text: "<See comments for Pick list>", options: { bold: false } },
        ],
        {
          y: footerTxtY,
          x: 9.33,
          h: footerH,
          w: 3.91,
          valign: "top",
          fontFace: "Segoe UI",
          fontSize: 11,
          color: "00B0F0",
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
  StoryPowerPoint: StoryPowerPoint,
};
