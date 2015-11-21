defmodule XLSXReader do
  use Application

  def start(_type, _args) do
    import Supervisor.Spec, warn: false

    children = [
      worker(XLSXReader.Parser, []),
    ]

    opts = [strategy: :one_for_one, name: XLSXReader.Supervisor]
    Supervisor.start_link(children, opts)
  end

  def parse(filename) do
    XLSXReader.Parser.parse(filename)
  end
end
