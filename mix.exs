defmodule XLSXReader.Mixfile do
  use Mix.Project

  def project do
    [app: :xlsx_reader,
     version: "0.0.1",
     elixir: "~> 1.1",
     build_embedded: Mix.env == :prod,
     start_permanent: Mix.env == :prod,
     compilers: [:javac] ++ Mix.compilers,
     deps: deps]
  end

  def application do
    [applications: [:logger],
     mod: {XLSXReader, []}]
  end

  defp deps do
    []
  end
end

defmodule Mix.Tasks.Compile.Javac do
  use Mix.Task

  @doc false
  def run(_args) do
    case :os.find_executable('javac') do
      [] ->
        throw(:javac_missing)
      javac ->
        {_, 0} = System.cmd(List.to_string(javac),
                            ["-cp", "priv/*",
                             "-d", "priv",
                             "java_src/XLSXReader.java"])
    end
  end
end
