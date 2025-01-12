// Controller to handle the home page
exports.getHomePage = (req, res) => {
    res.render('home', { title: 'Home Page', message: 'Welcome to the Express MVC App!' });
};
