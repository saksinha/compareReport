const express = require('express');
const multer = require('multer');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const { runAnalysis, FLOW_A_CONFIG, FLOW_B_CONFIG } = require('./analyzer.js');

const INTERNAL_PORT = 3000; // The port inside the container
const EXTERNAL_PORT = 3300; // The port you access from your browser

const app = express();

const upload = multer({ dest: 'uploads/' });

app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

app.post('/analyze', upload.fields([{ name: 'file1' }, { name: 'file2' }]), async (req, res) => {
    let uploadedFilePaths = [];
    try {
        const { isSingleFileMode, isWeeklyMode, isConciseMode, flowType } = req.body;
        const config = flowType === 'B' ? FLOW_B_CONFIG : FLOW_A_CONFIG;
        
        if (req.files.file1) uploadedFilePaths.push(req.files.file1[0].path);
        if (req.files.file2) uploadedFilePaths.push(req.files.file2[0].path);

        if (uploadedFilePaths.length === 0) {
            return res.status(400).json({ error: "No files were uploaded." });
        }

        const options = {
            isSingleFileMode: isSingleFileMode === 'true',
            isWeeklyMode: isWeeklyMode === 'true',
            isConciseMode: isConciseMode === 'true',
            filePaths: uploadedFilePaths,
            config: config
        };

        const reportContent = runAnalysis(options);
        
        res.send(reportContent);

    } catch (error) {
        console.error("Analysis Error:", error);
        res.status(500).json({ error: `An error occurred: ${error.message}` });
    } finally {
        uploadedFilePaths.forEach(filePath => {
            if (fs.existsSync(filePath)) {
                fs.unlinkSync(filePath);
            }
        });
    }
});

app.listen(INTERNAL_PORT, '0.0.0.0', () => {
    console.log(`Server is running!`);
    console.log(`Access the UI at http://localhost:${EXTERNAL_PORT}`);
});