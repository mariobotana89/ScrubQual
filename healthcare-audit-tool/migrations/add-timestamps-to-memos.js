// migrations/add-timestamps-to-memos.js

module.exports = {
    up: async (queryInterface, Sequelize) => {
      // Add columns without default values
      await queryInterface.addColumn('Memos', 'createdAt', {
        allowNull: true,
        type: Sequelize.DATE
      });
      await queryInterface.addColumn('Memos', 'updatedAt', {
        allowNull: true,
        type: Sequelize.DATE
      });
  
      // Populate the createdAt and updatedAt columns
      await queryInterface.sequelize.query(
        `UPDATE Memos SET createdAt = CURRENT_TIMESTAMP, updatedAt = CURRENT_TIMESTAMP`
      );
  
      // Set columns to not allow null values now that they are populated
      await queryInterface.changeColumn('Memos', 'createdAt', {
        allowNull: false,
        type: Sequelize.DATE,
        defaultValue: Sequelize.literal('CURRENT_TIMESTAMP')
      });
      await queryInterface.changeColumn('Memos', 'updatedAt', {
        allowNull: false,
        type: Sequelize.DATE,
        defaultValue: Sequelize.literal('CURRENT_TIMESTAMP')
      });
    },
  
    down: async (queryInterface, Sequelize) => {
      await queryInterface.removeColumn('Memos', 'createdAt');
      await queryInterface.removeColumn('Memos', 'updatedAt');
    }
  };
  