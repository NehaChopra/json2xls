require 'json2xls'

RSpec.describe Json2xls do
  it "has a version number" do
    expect(Json2xls::VERSION).not_to be nil
  end

  it "creates the worksheet on providing valid paths" do
    Json2xls::Generator.new(["../spec/sample.json","../spec/sample1.json"])
   end

  it "creates the worksheet on providing valid paths, worksheet name and path where to load the .xls file" do
    Json2xls::Generator.new(["../spec/sample.json","../spec/sample1.json"], {name: "Test", path: "#{ENV['HOME']}/Json2xls/"})
  end
end

