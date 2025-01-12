// Placeholder model (replace with database logic)
const users = [
    { id: 1, name: 'Alice', email: 'alice@example.com' },
    { id: 2, name: 'Bob', email: 'bob@example.com' },
];

// Get all users
exports.getAllUsers = () => {
    return users;
};

// Find user by ID
exports.findUserById = (id) => {
    return users.find(user => user.id === id);
};
