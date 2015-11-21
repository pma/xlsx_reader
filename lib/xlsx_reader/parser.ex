defmodule XLSXReader.Parser do
  use GenServer
  require Logger

  def start_link do
    GenServer.start_link(__MODULE__, [], [])
  end

  def init(_opts) do
    case :os.find_executable('java') do
      [] ->
        throw({:stop, :java_missing})
      java ->
        this_node = Atom.to_string(node())
        java_node =
          case String.split(Atom.to_string(node()), "@") do
            [name, server] ->
              String.to_atom(name <> "_java@" <> server);
            _Node ->
              throw({:bad_node_name, node()})
          end

        priv_dir = priv_dir(:xlsx_reader)
        classpath = priv_dir ++ '/*' ++ ':.'
        cwd = priv_dir

        port =
          :erlang.open_port({:spawn_executable, java},
                            [{:line, 1000}, :stderr_to_stdout,
                             {:args, ['-server',
                                      '-Xmx256m', '-Xms256m',
                                      '-XX:+UseConcMarkSweepGC', '-XX:+UseParNewGC',
                                      '-classpath', classpath,
                                      'XLSXReader',
                                      this_node, Atom.to_string(java_node),
                                      :erlang.get_cookie()]},
                             {:cd, cwd}])
        wait_for_ready(%{java_port: port, java_node: java_node})
    end
  end

  def handle_info({:nodedown, java_node}, %{java_node: java_node} = state) do
    {:stop, :nodedown, state}
  end

  def handle_info({port, {:data, {:eol, data}}}, %{java_port: port} = state) do
    Logger.debug List.to_string(data)
    {:noreply, state}
  end

  def wait_for_ready(%{java_port: port} = state) do
    receive do
      {^port, {:data, {:eol, 'READY'}}} ->
        true = :erlang.monitor_node(state.java_node, true)
        {:ok, state}
      info ->
        case handle_info(info, state) do
          {:noreply, new_state} ->
            wait_for_ready(new_state);
          {:stop, reason, _new_state} ->
            {:stop, reason}
        end
    end
  end

  def parse(filename) do
    send({:xlsx_reader_server, java_node()},
         {:parse, self(), String.to_char_list(filename)})
    receive do
      {:result, _pid, result} -> result
    after 30_000 ->
        throw(:timeout)
    end
  end

  def priv_dir(app) do
    case :code.priv_dir(app) do
      {:error, :bad_name} ->
        :error_logger.info_msg("Couldn't find priv dir for the application, using ./priv~n")
        "./priv"
      priv_dir -> :filename.absname(priv_dir)
    end
  end

  def java_node() do
    case String.split(Atom.to_string(node()), "@") do
      [name, server] ->
        String.to_atom(name <> "_java@" <> server);
      _Node ->
        throw({:bad_node_name, node()})
    end
  end
end
