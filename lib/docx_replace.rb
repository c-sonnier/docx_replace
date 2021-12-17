require 'docx_replace/version'
require 'zip'
require 'tempfile'

module DocxReplace
  class Doc
    attr_reader :document_content

    def initialize(path, temp_dir=nil)
      @zip_file = Zip::File.new(path)
      @temp_dir = temp_dir
      read_docx_file
    end

    def replace(pattern, replacement, multiple_occurrences=false)
      replace = replacement.to_s.encode(xml: :text)
      if multiple_occurrences
        @document_content.force_encoding('UTF-8').gsub!(pattern, replace)
      else
        @document_content.force_encoding('UTF-8').sub!(pattern, replace)
      end
    end

    def matches(pattern)
      @document_content.scan(pattern).map(&:first)
    end

    def unique_matches(pattern)
      matches(pattern)
    end

    alias uniq_matches unique_matches

    def commit(new_path=nil)
      write_back_to_file(new_path)
    end

    private

    DOCUMENT_FILE_PATH = 'word/document.xml'.freeze

    def read_docx_file
      @document_content = @zip_file.read(DOCUMENT_FILE_PATH)
    end

    def write_back_to_file(new_path=nil)
      buffer = Zip::OutputStream.write_buffer do |out|
        @zip_file.entries.each do |e|
          unless e.name == DOCUMENT_FILE_PATH
            out.put_next_entry(e.name)
            out.print e.get_input_stream.read
          end
        end
        out.put_next_entry(DOCUMENT_FILE_PATH)
        out.print @document_content
      end

      if new_path.nil?
        path = @zip_file.name
        FileUtils.rm(path)
      else
        path = new_path
      end

      File.open(path, 'wb') { |f| f.write(buffer.string) }
    end
  end
end
