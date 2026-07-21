const excelService = require("../services/excel.service");
const path = require("path");

exports.updateWorkbook = async (req, res) => {

    try {

        if (!req.file) {

            return res.status(400).json({

                success: false,

                message: "No Excel file uploaded."

            });

        }

        const parsedData = {

            date: req.body.date,

            dsf: req.body.dsf,

            todayValue: req.body.todayValue

        };

        console.log("Uploaded File:", req.file.path);

        console.log(parsedData);

        const result = await excelService.updateWorkbook(

            req.file.path,

            parsedData

        );

       const filename = path.basename(result.output);

return res.json({

    success: true,

    result: {
        ...result,
        filename,
        downloadUrl:
        `${req.protocol}://${req.get("host")}/api/excel/download/${filename}`
    }

});
    }

    catch (error) {

        console.error(error);

        return res.status(500).json({

            success: false,

            message: error.message

        });

    }

};
exports.downloadWorkbook = async (req, res) => {
    try {
        const { filename } = req.params;

        const filePath = path.join(
            __dirname,
            "..",
            "uploads",
            filename
        );

        return res.download(filePath);
    } catch (error) {
        console.error(error);

        return res.status(500).json({
            success: false,
            message: error.message,
        });
    }
};