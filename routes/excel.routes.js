const express = require("express");
const multer = require("multer");

const excelController = require("../controllers/excel.controller");

const router = express.Router();

const upload = multer({
    dest: "uploads/"
});

router.post(
    "/update",
    upload.single("file"),
    excelController.updateWorkbook
);

router.get(
    "/download/:filename",
    excelController.downloadWorkbook
);

router.get("/", (req, res) => {
    res.json({
        success: true,
        message: "Excel API Working"
    });
});

module.exports = router;