-module(xlerl).

-export([parse/1]).
-export([add_shared_string/2]).
-export([edit_cell/6]).
-export([render/2]).

-include_lib("xmerl/include/xmerl.hrl").

-define(worksheet_type, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet").

%- API
% Main objective is to load all XML content in simple_forms

% Gets a file binary
% Unpacks its zip and returns a map with the parsed content of each file.
% Binary files are left untouched.
% #{InternalFilename => ParsedContent}
parse({_Filename, _Binary} = XlsxFile) ->
    zip:foldl(fun xml_parse/4, #{}, XlsxFile).

add_shared_string(String, Xlsx) ->
    SharedStringsFile = "xl/sharedStrings.xml",
    #{SharedStringsFile := SharedStrings} = Xlsx,
    #xmlElement{name = sst,
                attributes = Attributes,
                content = Strings} = SharedStrings,
    [#xmlAttribute{name = xmlns, value = _NS} = NameSpace,
     #xmlAttribute{name = count, value = Count} = CountAttr,
     #xmlAttribute{name = uniqueCount, value = UniqueCount} = UniqueCountAttr
    ] = Attributes,
    Pos = list_to_integer(UniqueCount) + 1,
    NewStrings = Strings ++ [unique_string_element(String, Pos)],
    NewAttributes = [
        NameSpace,
        CountAttr#xmlAttribute{value =
                            integer_to_list(list_to_integer(Count) + 1)},
        UniqueCountAttr#xmlAttribute{value = integer_to_list(Pos)}
    ],
    NewSharedStrings = SharedStrings#xmlElement{
        attributes = NewAttributes,
        content = NewStrings},
    {Pos, Xlsx#{SharedStringsFile => NewSharedStrings}}.

edit_cell(SheetName, Column, Row, DataType, Value, Xlsx) ->
    % first we find the file...
    SheetFile = lookup_sheet_file(SheetName, Xlsx),
    % io:format("Sheet file =  ~p\n", [SheetFile]),
    #{SheetFile := SheetXml} = Xlsx,
    % SheetXml2 = edit_sheet_cell(Column, Row, DataType, Value, SheetXml),
    % Xlsx#{SheetFile := SheetXml2}.
    Xlsx.

render(Filename, Xlsx) ->
    ZipBinaries = #{K => export_xml(V) || K := V <- Xlsx},
    zip:create(Filename, maps:to_list(ZipBinaries), [memory]).

%- Internal --------------------------------------------------------------------

lookup_sheet_file(SheetName, Xlsx) ->
    #{"xl/workbook.xml" := WorkbookXml} = Xlsx, % TODO lookup workbook file
    [Sheet] = [S || S <- workbook_sheets(WorkbookXml), xml_elem_has_attr('name',
                                                                  SheetName, S)],
    #xmlElement{name = 'sheet', attributes = Attributes} = Sheet,
    [#xmlAttribute{name = 'name', value = SheetName},
     #xmlAttribute{name = 'sheetId', value = _},
     #xmlAttribute{name = 'r:id', value = RelId}] = Attributes,
    {?worksheet_type, Target} = lookup_relationship(RelId, "xl/workbook.xml", Xlsx),
    filename:join("xl", Target).

workbook_sheets(#xmlElement{name = 'workbook', content = Content}) ->
    [SheetsList] = [C || #xmlElement{name = 'sheets',
                                     content = C} <- Content],
    SheetsList.

lookup_relationship(Id, Filename, Xlsx) ->
    RelFile = filename:basename(Filename) ++ ".rels",
    RelPath = filename:join([filename:dirname(Filename), "_rels", RelFile]),
    #{RelPath := RelXml} = Xlsx,
    #xmlElement{name = 'Relationships', content = RelationShips} = RelXml,
    [R] = [R || R <- RelationShips, xml_elem_has_attr('Id', Id, R)],
    #xmlElement{name = 'Relationship',
                attributes = Attributes} = R,
    [#xmlAttribute{name = 'Id', value = Id},
     #xmlAttribute{name = 'Type', value = Type},
     #xmlAttribute{name = 'Target', value = Target}] = Attributes,
    {Type, Target}.

xml_elem_has_attr(Name, Value, #xmlElement{attributes = Attributes}) ->
    lists:any(fun(A) -> match_attr(Name, Value, A) end, Attributes).

match_attr(Name, Value, #xmlAttribute{name = Name, value = Value}) -> true;
match_attr(_, _, _) -> false.

export_xml(Bin) when is_binary(Bin) -> Bin;
export_xml(#xmlElement{} = E) ->
    % FL = lists:flatten(xmerl:export_simple([E], xmerl_xml)),
    unicode:characters_to_binary(xmerl:export_simple([E], xmerl_xml)).
    % % CursedChars = [I || I <- XML, (I < 0) or (I > 255)],
    % % io:format("Chars= ~p\n", [CursedChars]),
    % case file:write_file(N, XML) of
    %     ok -> ok;
    %     Error -> io:format("Error ~p, writing ~p\n",[Error, N])
    % end,

% xml_parse("xl/styles.xml", _GetInfo, GetBin, Store) ->
%     Store#{"xl/styles.xml" => GetBin()}; % file is problematic
xml_parse(Filename, _GetInfo, GetBin, Store) ->
    Content = case filename:extension(Filename) of
        Xml when Xml == ".xml" ; Xml == ".rels" ->
            {XmlElement, []} = xmerl_scan:string(binary_to_list(GetBin())),
            XmlElement;
        _ -> GetBin()
    end,
    Store#{Filename => Content}.

unique_string_element(String, Pos) ->
    #xmlElement{
        name = si,
        expanded_name = si,
        nsinfo = [],
        namespace = #xmlNamespace{},
        parents = [{sst, 1}],
        pos = Pos,
        content = [
            #xmlElement{
                name = t,
                expanded_name = t,
                namespace = #xmlNamespace{},
                parents = [{si, Pos}, {sst, 1}],
                pos = 1,
                content = [#xmlText{
                    parents = [{t, 1}, {si, Pos}, {sst, 1}],
                    pos = 1,
                    value = String,
                    type = text
                }]
            }
        ]
    }.
