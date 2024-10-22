// config/database.js
const { Sequelize, DataTypes } = require('sequelize');
const sequelize = new Sequelize({
    dialect: 'sqlite',
    storage: 'database.sqlite'
});

// Define Memo model
const Memo = sequelize.define('Memo', {
    content: {
        type: DataTypes.TEXT,
        allowNull: false
    },
    classification: {
        type: DataTypes.ENUM('informational', 'actionable'),
        allowNull: false
    },
    status: {
        type: DataTypes.ENUM('routed', 'pending'),
        allowNull: false
    },
    assigned_department: {
        type: DataTypes.TEXT,
        allowNull: true
    }
}, {
    timestamps: true,
    createdAt: 'created_at',
    updatedAt: false
});

// Define Feedback model
const Feedback = sequelize.define('Feedback', {
    memo_id: {
        type: DataTypes.INTEGER,
        references: {
            model: Memo,
            key: 'id'
        }
    },
    user_classification: {
        type: DataTypes.ENUM('informational', 'actionable'),
        allowNull: false
    },
    correct_classification: {
        type: DataTypes.ENUM('informational', 'actionable'),
        allowNull: true
    }
}, {
    timestamps: true,
    createdAt: 'created_at',
    updatedAt: false
});

// Sync database
(async () => {
    await sequelize.sync({ force: true });
    console.log("Database synced!");
})();

module.exports = { Memo, Feedback };
