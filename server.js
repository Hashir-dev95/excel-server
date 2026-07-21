require("dotenv").config();

const express = require("express");
const cors = require("cors");

const excelRoutes = require("./routes/excel.routes");

const app = express();

app.use(cors());
app.use(express.json());

// Routes
app.use("/api/excel", excelRoutes);

// Test Route
app.get("/", (req, res) => {
    res.send("Excel Server is Running");
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});