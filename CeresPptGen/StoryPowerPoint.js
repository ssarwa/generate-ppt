
let PptxGenJS = require("pptxgenjs");

class StoryPowerPoint {

    constructor(projects) {
        this.projects = projects;
        this.pptx = new PptxGenJS();
        this.masterPlaceholder = 'PLACEHOLDER_SLIDE';
        
        //Pseudo constants
        this.COLOR_WHITE = 'FFFFFF';
        this.COLOR_BLACK = '000000';
        this.COLOR_RED = 'fc0303';
        this.COLOR_GREEN = '00b04f';
        this.COLOR_BLUE = '4472c4';
        this.COLOR_DARK_BLUE = '44546a';
        this.COLOR_AMBER = 'ffbf00';
        this.COLOR_GRAY = '808080';
        this.COLOR_LIGHT_GRAY = 'f2f2f2';
        this.COLOR_NOTES_HEADER = '0078D4'; 
        this.COLOR_BOX_BACKGROUND = '3B3B3B';
        this.COLOR_CELL_HEADER = this.COLOR_WHITE; //'4473c4';
        this.COLOR_CELL_DARK = this.COLOR_LIGHT_GRAY //'cfd5ea';
        this.COLOR_CELL_LIGHT = this.COLOR_WHITE // 'e9ebf5';
        
        this.tableHeaderOptions = { fill:this.COLOR_CELL_HEADER, fontFace: 'Calibri' , fontSize: 8 , align: 'center' , bold: true};
        this.tableDarkOptions = { fill:this.COLOR_CELL_DARK, fontFace: 'Calibri' , fontSize: 8 , align: 'left' , bold: false};
        this.tableLightOptions = { fill:this.COLOR_CELL_LIGHT, fontFace: 'Calibri' , fontSize: 8 , align: 'left' , bold: false};
        this.tableRowH = 0.32;
        this.tableColW = 1.4;


        var today = new Date().toISOString(); 
        today = today.replace(new RegExp(":", "g"),"-");
        this.filename = "FTA-CustomerStories-" + today + ".pptx";
    }

    generate() {
        this.pptx.layout = 'LAYOUT_WIDE';
        this.addMaster();
       
        this.projects.forEach(project => 
        {
            var slide = this.pptx.addSlide(this.masterPlaceholder);
            this.addTitleBar(slide, project);
            this.addNotesHeaders(slide, project);
            this.addNotesBodies(slide, project);
            this.addFooters(slide, project); 
        }
        );

    }

    addMaster() {
        var imageY = 6.7;
        var imageH = 0.6;
        var imageW = 0.6;
        this.pptx.defineSlideMaster({
            title: this.masterPlaceholder,
            bkgd:  this.COLOR_BLACK,
            /*objects: [
                {'rect': { fill: '96F0FF', 
                    y: 6.6, x: '1%', w: '98%', h: 0.8 }},
                {'image': { path:'img/country.jpg', 
                    y: imageY, x: "15%", w: imageW, h: imageH }},
                {'image': { path:'img/industry.jpg', 
                    y: imageY, x: "35%", w: imageW, h: imageH }},
                {'image': { path:'img/segment.jpg', 
                    fill:this.COLOR_WHITE,
                    y: imageY,x: "55%", w: imageW, h: imageH}},
                {'image': { path:'img/duration.jpg', 
                    y: imageY, x: "75%", w: imageW, h: imageH
                }}
            ],*/
          });
        }
    
    addTitleBar(slide, project){
       
        slide.addText("Logo", 
        { 
            shape:this.pptx.shapes.RECTANGLE,
            y: 0.14, x: 0.17, w: 1, h: 0.8, 
            fill: this.COLOR_BOX_BACKGROUND, color: this.COLOR_WHITE,
            fontFace: 'Calibri' , fontSize: 24 , align: 'Center' 
        });
        slide.addText(project.NominationCustomerName, 
        { 
            shape:this.pptx.shapes.RECTANGLE,
            y: 0.14, x: 1.17, w: 6.5, h: 0.8, 
            color: this.COLOR_WHITE,
            fontFace: 'Calibri' , fontSize: 24 , align: 'left' 
        });
        slide.addText("Customer quote here", 
        { 
            shape:this.pptx.shapes.RECTANGLE,
            y: 0.14, x: 6.95, w: 4.88, h: 0.8, 
            color: this.COLOR_WHITE,
            fontFace: 'Calibri' , fontSize: 14 , align: 'right', italic: true 
        });
    }

   
    addNotesHeaders(slide, project) {
        slide.addText('Customer Overview and Objective​', 
        { 
            shape:this.pptx.shapes.RECTANGLE,
            autoFit: false, y: 1, x: 0.1724727, w: 4.242982, h: 0.5, 
            fill: this.COLOR_NOTES_HEADER, color: this.COLOR_WHITE, 
            fontFace: 'Calibri' , fontSize: 11 , align: 'center' , bold: true
        });
        slide.addText('Business Impact & Learnings', 
        { 
            shape:this.pptx.shapes.RECTANGLE,
            autoFit: false, y: 1, x: 4.497233, w: 4.355464, h: 0.5, 
            fill: this.COLOR_NOTES_HEADER, color: this.COLOR_WHITE, 
            fontFace: 'Calibri' , fontSize: 11 , align: 'center' , bold: true
        });
        
        slide.addText('How did FTA support and guide the customer?', 
        { 
            shape:this.pptx.shapes.RECTANGLE,
            autoFit: false, y: 1, x: 8.924829, w: 4, h: 0.5, 
            fill: this.COLOR_NOTES_HEADER, color: this.COLOR_WHITE, 
            fontFace: 'Calibri' , fontSize: 11 , align: 'center' , bold: true 
        });
    }
    
    dateFromUtc(utc) {
        //Preconditions 
        if(!utc || utc.length  < 13 || 
                (!Number.isInteger(parseInt(utc.substr(0,4))) ) ){
            //console.log("DateUtil.UTCtoMedium: utc parameter invalid: " + utc);
            return "";
        }
        //avoid "0001-01-01T08:00:00+00:00"
        if (utc.substr(0,4) < "2000") return "";
    
        var year = utc.substr(0,4);
        var month = parseInt(utc.substr(5,2)) - 1; //Months are 0-11
        var day = utc.substr(8,2);
        var hours = parseInt(utc.substr(11,2));
        if(hours > 11)
            day++; //goes to next day, if it's more than days in the month, date object adjusts accordingly
        var dateObj = new Date(year,month,day);
        return dateObj;
    }

    daysBetween ( date1, date2 ) {
        //Get 1 day in milliseconds
        var one_day=1000*60*60*24;
        // Convert both dates to milliseconds
        var date1_ms = date1.getTime();
        var date2_ms = date2.getTime();
      
        // Calculate the difference in milliseconds
        var difference_ms = date2_ms - date1_ms;
          
        // Convert back to days and return
        return Math.round(difference_ms/one_day); 
      }

    addFooters(slide, project) {
        var imageY = 6.6;
        var imageH = 0.86;
        var imageW = 2.2;
        var start = this.dateFromUtc(project.StartDate);
        var end = this.dateFromUtc(project.EndDate);
        var duration = this.daysBetween(start,end);

        slide.addText("pic", 
        { 
            shape:this.pptx.shapes.RECTANGLE,
            y: imageY, x: 0.17, h: imageH, w: 1, valign: 'top', 
            autoFit: false, fill: this.COLOR_BOX_BACKGROUND, color: this.COLOR_BLUE, 
            fontFace: 'Segoe UI' , fontSize: 11 , align: 'left' , bold: true 
        });
        slide.addText("Customer Location:\n" + project.PhysicalLocation, 
        { 
            shape:this.pptx.shapes.RECTANGLE,
            y: imageY, x: 1.17, h: imageH, w:imageW, valign: 'top', 
            autoFit: false, fill: this.COLOR_BOX_BACKGROUND, color: this.COLOR_BLUE, 
            fontFace: 'Segoe UI' , fontSize: 11 , align: 'left' , bold: true 
        });

        slide.addText("pic", 
        { 
            shape:this.pptx.shapes.RECTANGLE,
            y: imageY, x: 3.37, h: imageH, w: 1, valign: 'top', 
            autoFit: false, fill: this.COLOR_BOX_BACKGROUND, color: this.COLOR_BLUE, 
            fontFace: 'Segoe UI' , fontSize: 11 , align: 'left' , bold: true 
        });
        slide.addText("Customer Solution: " + project.QualifiedIndustry + "\nProject Duration: " + duration + " Days", 
        { 
            shape:this.pptx.shapes.RECTANGLE,
            y: imageY, x: 4.36, h: imageH, w: 2.68, valign: 'top', 
            autoFit: false, fill: this.COLOR_BOX_BACKGROUND, color: this.COLOR_BLUE, 
            fontFace: 'Segoe UI' , fontSize: 11 , align: 'left' , bold: true 
        });
        
        slide.addText("pic", 
        { 
            shape:this.pptx.shapes.RECTANGLE,
            y: imageY, x: 7.04, h: imageH, w: 1, valign: 'top', 
            autoFit: false, fill: this.COLOR_BOX_BACKGROUND, color: this.COLOR_BLUE, 
            fontFace: 'Segoe UI' , fontSize: 11 , align: 'left' , bold: true 
        });
        slide.addText("Customer Industry:\n" + project.QualifiedIndustry,
        { 
            shape:this.pptx.shapes.RECTANGLE,
            y: imageY, x: 8.04, h: imageH, w: 2.09, valign: 'top', 
            autoFit: false, fill: this.COLOR_BOX_BACKGROUND, color: this.COLOR_BLUE, 
            fontFace: 'Segoe UI' , fontSize: 11 , align: 'left' , bold: true 
        });
        
        slide.addText("pic", 
        { 
            shape:this.pptx.shapes.RECTANGLE,
            y: imageY, x: 10.13, h: imageH, w: 1, valign: 'top', 
            autoFit: false, fill: this.COLOR_BOX_BACKGROUND, color: this.COLOR_BLUE, 
            fontFace: 'Segoe UI' , fontSize: 11 , align: 'left' , bold: true 
        });
        slide.addText("Customer Segment:\n" + project.QualifiedSegment, 
        { 
            shape:this.pptx.shapes.RECTANGLE,
            y: imageY, x: 11.13, h: imageH, w:1.79, valign: 'top', 
            autoFit: false, fill: this.COLOR_BOX_BACKGROUND, color: this.COLOR_BLUE, 
            fontFace: 'Segoe UI' , fontSize: 11 , align: 'left' , bold: true 
        });
    }

//    findActivityNoteText(project,title) {
//        //solution
//        title = title.toLowerCase();
//        var noteText = "";
//        for (let i = 0; i < project.activities.length; i++) {
//            if(project.activities[i].ActivityName.toLowerCase().search(title) > -1) {
//                noteText = project.activities[i].PlainNoteText; //previously calculated
//                 break; 
//             }
//        }
//        return noteText;
//   }

    addNotesBodies(slide, project) {

        var top = 1.5;
        var height = 5;
        /*AddLink
        slide.addText(
            [{
                text: 'Open in Ceres...',
                options: { hyperlink:{ url:project.Link, tooltip:'Open project in Ceres' } }
            }],
            { y: 7, x: 0.1, w: 2.935527, h: 0.2827406, 
            fontFace: 'Calibri' , fontSize: 12 , align: 'left'}
        )*/

        //Project Description
        slide.addText("Project Description", 
        { 
            shape:this.pptx.shapes.RECTANGLE, 
            autoFit: false, y: top, x: 0.1724727, w: 4.229456, h: 3,shrinkText:true,
            valign: 'top', fill: this.COLOR_BOX_BACKGROUND, color: this.COLOR_WHITE,
            fontFace: 'Calibri' , fontSize: 10 , align: 'left' , bold: false
        });

        //impact, learnings
        slide.addText("Business Impact:\n" + "learnings"
        + "Learnings:\n" + "learnings", 
        { 
            shape:this.pptx.shapes.RECTANGLE,
            autoFit: false, y: top, x: 4.497233, w: 4.355464, h: height,shrinkText:true, 
            valign: 'top', fill: this.COLOR_BOX_BACKGROUND, color: this.COLOR_WHITE,
            fontFace: 'Calibri' , fontSize: 10 , align: 'left' , bold: false
        });

        slide.addText("Solution Summary:\n" + "Solution" ,
            { 
                shape:this.pptx.shapes.RECTANGLE,
                autoFit: false, y: top, x:8.924829, w: 4, h: height, shrinkText:true,
                valign: 'top',  fill: this.COLOR_BOX_BACKGROUND, color: this.COLOR_WHITE,
                fontFace: 'Calibri' , fontSize: 10 , align: 'left' , bold: false
        });

        slide.addText("Collaborators​\n", 
        { 
            shape:this.pptx.shapes.RECTANGLE, 
            autoFit: false, y: 4.6, x: 0.1724727, w: 4.229456, h: 0.4,
            valign: 'top', fill: this.COLOR_BOX_BACKGROUND, color: this.COLOR_BLUE,
            fontFace: 'Calibri', fontSize: 14, align: 'left', bold: true, italic: true
        });

        slide.addText("FastTrack for Azure: " + "value" + "\n" + "Field: " + "value" + "\n" + "EPM: " + "value" + "\n" + "Skilling: " + "value" + "\n" + "Partner: " + "value" + "\n", 
        { 
            shape:this.pptx.shapes.RECTANGLE, 
            autoFit: false, y: 5, x: 0.1724727, w: 4.229456, h: 1.5,shrinkText:true,
            valign: 'top', fill: this.COLOR_BOX_BACKGROUND, color: this.COLOR_WHITE,
            fontFace: 'Calibri', fontSize: 10, align: 'left', bold: false
        });

    }
}

module.exports = {
    StoryPowerPoint : StoryPowerPoint
}