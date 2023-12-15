-module(xlerl).

-export([parse/1]).
-export([render/1]).

%- API

% Main objective is to load all XML content in simple_forms
%
% Alternative:
% Use sax_parsing over simple_form?
% How to store data from a sax parser such that is easely assessible and
% friendly for xmerl to convert back to XML?
%

% Gets a file binary
% Unpacks its zip and returns a map with the parsed content of each file
% #{InternalFilename => ParsedContent}
parse({_Filename, _Binary} = XlsxFile) ->
    zip:foldl(fun xml_parse/4, #{}, XlsxFile).

render(Xlsx) ->
    <<"LOL">>.

%- Internal --------------------------------------------------------------------

xml_parse(Filename, _GetInfo, _GetBin, Store) ->
    % io:format("~p\n",[Filename]),
    Term = {},
    Store#{Filename => Term}.
