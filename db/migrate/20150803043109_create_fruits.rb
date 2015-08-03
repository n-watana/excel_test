class CreateFruits < ActiveRecord::Migration
  def change
    create_table :fruits do |t|
      t.string :name
      t.integer :season_id

      t.timestamps null: false
    end
  end
end
