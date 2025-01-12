const express = require('express');
const router = express.Router();
const homeController = require('../contorller/homeController');

// Home route
router.get('/', homeController.getHomePage);

module.exports = router;
