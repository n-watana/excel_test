json.array!(@fruits) do |fruit|
  json.extract! fruit, :id, :name, :season_id
  json.url fruit_url(fruit, format: :json)
end
