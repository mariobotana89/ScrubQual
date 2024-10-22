'use strict';
module.exports = {
  up: async (queryInterface, Sequelize) => {
    const tableDefinition = await queryInterface.describeTable('Users');
    
    // Add columns only if they don't exist
    if (!tableDefinition.isApproved) {
      await queryInterface.addColumn('Users', 'isApproved', {
        type: Sequelize.BOOLEAN,
        defaultValue: false
      });
    }
    if (!tableDefinition.resetToken) {
      await queryInterface.addColumn('Users', 'resetToken', {
        type: Sequelize.STRING,
        allowNull: true
      });
    }
    if (!tableDefinition.resetTokenExpiration) {
      await queryInterface.addColumn('Users', 'resetTokenExpiration', {
        type: Sequelize.DATE,
        allowNull: true
      });
    }
  },
  down: async (queryInterface, Sequelize) => {
    const tableDefinition = await queryInterface.describeTable('Users');
    
    // Remove columns if they exist
    if (tableDefinition.isApproved) {
      await queryInterface.removeColumn('Users', 'isApproved');
    }
    if (tableDefinition.resetToken) {
      await queryInterface.removeColumn('Users', 'resetToken');
    }
    if (tableDefinition.resetTokenExpiration) {
      await queryInterface.removeColumn('Users', 'resetTokenExpiration');
    }
  }
};
