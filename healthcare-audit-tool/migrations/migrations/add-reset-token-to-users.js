'use strict';
module.exports = {
  up: async (queryInterface, Sequelize) => {
    const tableDefinition = await queryInterface.describeTable('Users');
    
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
    
    if (tableDefinition.resetToken) {
      await queryInterface.removeColumn('Users', 'resetToken');
    }
    if (tableDefinition.resetTokenExpiration) {
      await queryInterface.removeColumn('Users', 'resetTokenExpiration');
    }
  }
};
