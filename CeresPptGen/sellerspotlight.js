let PptxGenJS = require("pptxgenjs");
var dateFormat = require("dateformat");

class sellerspotlight {
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
    this.filename = project.ProjectCustomerName.replace(" ", "-") + "-SellerSpotlight" + "-" + today + ".pptx";
  
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
      '**Intended Audience** - Pillar reviews, Account Teams, FTA LT – This is to be part of our internal Flipbook\nThis slide should ALWAYS be in place for any customer story. The following Technical slides may or may not be required, but this Business Impact Overview is always the starting point for any customer story being triggered.\n\nThe Technical slides to follow may be used to add context for various technical audiences. Several combinations of the Technical slides may be used to cater your story to the target audience. If the PG prefers to see buckets with challenges and associated learnings and impact, the final slide is this PPT may be ideal. If some prefers a higher-lever technical overview, then the first technical slide option would be better suited.\n\nCustomer Overview and Objective\nThis section needs to outline who the customer is, the customer’s objective, the "WHY" .. Why were they pursuing this change, and what was the ideal outcome? Tell a brief story around the Why and What – you should be able to pull a lot of this content from the Project Overview in Ceres or your scoping notes prior to project creation\n\nAzure Solution Overview\nWork with your FTA Engineer to populate this section if you are unsure. Provide a high-level overview of the solution we helped them implement, as well as key services and components of the FTA delivery (ie, Governance discussions, Arch Design Review, over-the-shoulder working sessions, etc…)\n\nBusiness Results and Impact\nThis is one of the most critical pieces of the Business Impact Overview. Most audiences viewing this are really going to want to know what the Customer Impact was and what it has enabled the customer to do next. Be sure to discuss FTA\'s impact in the customer achieving their desired outcomes. There should always be something to tell here in terms of “What’s Next” – maybe they are moving additional machines to Azure, planning modernization efforts over the next 6 months, evaluating supplemental services, etc…'
    );
  }

  addMaster() {
    this.pptx.defineSlideMaster({
      title: this.masterPlaceholder,
      bkgd: this.COLOR_BLACK,
    });
  }

  addTitleBar(slide, project) {
    slide.addText("Picture", {
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

    slide.addText(
        [
            {
              text: "Seller Name + FastTrack for Azure\n",
              options: { fontSize: 22, },
            },        {
              text: "Role - Org",
              options: { fontSize: 18, },
            },
          ], 
     {
      shape: this.pptx.shapes.RECTANGLE,
      autoFit: true,
      y: 0.13,
      x: 1.3,
      w: 5.4,
      h: 0.6,
      color: this.COLOR_WHITE,
      fontFace: "Segoe UI Semibold",
      align: "left",
      valign: "top",
    });

	slide.addShape(this.pptx.shapes.LINE, {
		x: 7.63,
		y: 0.13,
		w: 0.0,
		h: 0.53,
		line: { color: "FFFFFF", width: 2, },
	});

    slide.addImage({
            path: "img/quotes.png",
            y: 0.18,
            x: 7.8,
            w: 0.35,
            h: 0.34,
          });
      
    slide.addText("“Insert seller quote about what you would tell peers about FTA”",
     {
      shape: this.pptx.shapes.RECTANGLE,
      y: 0.21,
      x: 8.26,
      w: 6,
      h: .2,
      color: this.COLOR_WHITE,
      fontFace: "Segoe UI",
      fontSize: 11,
      italic: true,
    });
  }

  addNotesHeaders(slide, project) {
    slide.addText(" Opportunity", {
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

    slide.addText(" FastTrack for Azure Experience", {
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

    //Opportunity

    slide.addText(
      [
        {
            text: "Customer Name: ",
            options: { bold: true, bullet: false, breakLine: false, fontSize: 11 },
        },
        {
            text: project.ProjectCustomerName,
            options: { bullet: false, breakLine: true, fontSize: 11 },
        },
        {
            text: "Customer Background:",
            options: { bold: true, bullet: false, breakLine: true, fontSize: 11 },
        },
        {
            text: "Who are they?  What is their business?  Industry challenges?",
            options: { bullet: false, breakLine: true, fontSize: 11 },
        },
        {
            text: "Customer Opportunity:",
            options: { bold: true, bullet: false, breakLine: true, fontSize: 11 },
        },
        {
          text: "What were they doing and why?  (Project overview, business drivers, end user, etc…)",
          options: { bullet: { indent: 10 } },
        },
        {
          text: "What is the timeline and why is it important?",
          options: { bullet: { indent: 10 } },
        },
        {
          text: "What concerns do they have for their business?  Their users?  The solution?",
          options: { bullet: { indent: 10 } },
        },
      ],
      {
        y: top,
        x: 0.12,
        w: 3.49,
        h: 5.2,
        shrinkText: true,
        fill: this.COLOR_BOX_BACKGROUND,
        color: this.COLOR_WHITE,
        fontFace: "Segoe UI",
        fontSize: 11,
        valign: "top",
      }
    );

        slide.addText("Customer\nLogo", {
          shape: this.pptx.shapes.ROUNDED_RECTANGLE,
          y: 1.53,
          x: 5.53,
          w: 1.17,
          h: 0.77,
          color: this.COLOR_WHITE,
          fontFace: "Segoe UI",
          fontSize: 14,
          align: "center",
          fill: { color: this.COLOR_BOX_BACKGROUND },
          line: { color: "FFFFFF", width: 1 },
        });
    
    slide.addImage({
      path: "img/acrchart.png",
      y: 2.57,
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
        y: 2.93,
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
            text: "FTA Positioning",
            options: { bold: true, breakLine: true, fontSize: 12 },
        },
        {
            text: "  insert seller response here\n",
            options: { italic: true, breakLine: true, fontSize: 11, },
        },
        {
            text: "Impact",
            options: { bold: true, breakLine: true, fontSize: 12 },
        },
        {
            text: "  insert seller response here\n",
            options: { italic: true, breakLine: true, fontSize: 11,  },
        },
        {
            text: "What would you tell your peers about FTA?",
            options: { bold: true, breakLine: true, fontSize: 12 },
        },
        {
            text: "  insert seller response here\n",
            options: { italic: true, breakLine: true, fontSize: 11,  },
        },
        {
            text: "What’s Next?",
            options: { bold: true, breakLine: true, fontSize: 12 },
        },
        {
            text: "  insert seller response here\n",
            options: { italic: true, breakLine: true, fontSize: 11,  },
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
      });
   }


    // footer

  addFooters(slide, project) {
    var footerY = 6.7;
    var footerTxtY = 6.72;
    var footerH = 0.6;
    var imageH = 0.5;
    var imageW = 0.5;

    var startDate = dateFormat(project.StartDate, "mmm-yyyy");
    var endDate = dateFormat(project.EndDate, "mmm-yyyy");

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
        { text: startDate + " to " + endDate, options: { bold: false } },
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
}

module.exports = {
  sellerspotlight: sellerspotlight,
};
