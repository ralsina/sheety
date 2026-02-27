require "baked_file_system"

module Sheety
  # Embedded file system for generated template files
  # These files are baked into the sheety binary at compile time
  class EmbeddedFiles
    extend BakedFileSystem

    bake_folder "../../src"
  end
end
