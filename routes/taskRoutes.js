const express = require('express');
const multer = require('multer');
const taskController = require('../contorller/taskController');

const router = express.Router();

// File upload setup
const upload = multer({ dest: 'uploads/' });

router.get('/', taskController.index);
router.get('/upload', taskController.uploadForm);
// router.post('/convert', upload.single('textfile'), taskController.converTextToExcel);
router.post("/converts", upload.array("textfile", 10), taskController.converTextToExcelWithMutiple)
module.exports = router;
