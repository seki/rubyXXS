require 'nokogiri'
require 'tmpdir'
require 'fileutils'

class RubyXXS
  def initialize(fname)
    @fname = fname
    @dir = Dir.mktmpdir('XXS', '/tmp')
    extract
  end
  attr_reader :dir

  def extract
    system("unzip #{@fname} -d #{@dir}")
  end

  def archive(dest)
    system("cd #{@dir}; zip -r #{dest} *")
  end

  def close
    FileUtils.remove_entry(@dir) if @dir
  ensure
    @dir = nil
  end

  def load_sheet(n)
    fname = @dir + "/xl/worksheets/sheet#{n}.xml"
    File.open(fname) {|fp| Nokogiri::XML(fp)}
  end

  def edit_sheet(n)
    fname = @dir + "/xl/worksheets/sheet#{n}.xml"
    doc = File.open(fname) {|fp| Nokogiri::XML(fp)}
    yield(doc)
    File.open(fname, "w") {|fp| doc.write_to(fp)}
  end
end

if __FILE__ == $0

  xxs = RubyXXS.new(ARGV.shift)

  (1..6).each do |n|
    xxs.edit_sheet(n) do |doc|
      c = doc.xpath(
        # '//ns:c[@r="G5"]',
        "//ns:c/ns:v[contains(text(), '44275')]",
        {ns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
      ).first
      next unless c
      pp c
      pp c.parent
    end
  end
  #xxs.archive('/tmp/new_xxs.xlsx')
  xxs.close
end