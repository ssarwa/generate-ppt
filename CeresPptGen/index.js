const pptx = require("./pptx");

module.exports = async function (context, req) {
    context.log('CeresPptGen function processed a request.\nBody: ' + JSON.stringify(req.body));

    try {
        //if(req.body.option == "download")
        await pptx.generatePptxDownload(context, req);
    //else
    //    await pptx.generatePptxSAS(context, req);
    } catch (err) {
        console.log(err.message);
        return res.status(500).send({message: `Request Failed: ${err} `})
    }


}