%%%-------------------------------------------------------------------
%% @doc xlerl public API
%% @end
%%%-------------------------------------------------------------------

-module(xlerl_app).

-behaviour(application).

-export([start/2, stop/1]).

start(_StartType, _StartArgs) ->
    xlerl_sup:start_link().

stop(_State) ->
    ok.

%% internal functions
