% @doc API to parse, edit and write xlsx files with xmerl and zip
-module(xlerl).

% API
-export([parse/1]).
-export([add_shared_string/2]).
-export([write/5]).
-export([read/4]).
-export([render/2]).

% Includes
-include_lib("xmerl/include/xmerl.hrl").

%--- Macros --------------------------------------------------------------------

-define(worksheet_type, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet").

%--- API -----------------------------------------------------------------------

% @doc Gets a file binary
% Unpacks its zip and returns a map with the parsed content of each file.
% Binary files are left untouched.
% #{InternalFilename => ParsedContent}
-spec parse({Filename :: file:name(), Binary :: binary()}) ->
    {ok, Xlsx :: map()} | {error, Reason :: term()}.
parse({_Filename, _Binary} = XlsxFile) ->
    zip:foldl(fun xml_parse/4, #{}, XlsxFile).

% @doc Adds a string to the xl/sharedStrings.xml file
% The new string is inserted at the end of the list and the list length is returned.
% This should allow the user to just insert the string ID in other cells.
% Currently, this method is not working properly but may be fixed in future.
% It could also be embedded in the write call as an optimization.
-spec add_shared_string(String :: binary(), Xlsx :: map()) ->
    {StringIndex :: string(), NewXlsx :: map()}.
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
    {UniqueCount, Xlsx#{SharedStringsFile => NewSharedStrings}}.

% @doc Extract a single value in a precise locations on a selected sheet
% Returns 'empty' if cell is empty
-spec read(SheetName :: string(), Column :: string(),
           Row :: string(),Xlsx :: map()) ->
    Value :: string().
read(SheetName, Column, Row, Xlsx) ->
    SheetFile = lookup_sheet_file(SheetName, Xlsx),
    #{SheetFile := SheetXml} = Xlsx,
    ColumnName = Column ++ Row, % e.g. "B1" = "B" ++ "1"
    read_worksheet(Row, ColumnName, SheetXml).

% @doc Directly inserts values in precise locations on a selected sheet
% For now it expects sheets with unique names,
% which might not always be the case.
-spec write(SheetName :: string(), Column :: string(),
            Row :: string(), Value :: term(), Xlsx :: map()) ->
    NewXlsx :: map().
write(SheetName, Column, Row, Value, Xlsx) ->
    SheetFile = lookup_sheet_file(SheetName, Xlsx),
    #{SheetFile := SheetXml} = Xlsx,
    ColumnName = Column ++ Row, % e.g. "B1" = "B" ++ "1"
    SheetXml2 = write_worksheet(Row, ColumnName, Value, SheetXml),
    Xlsx#{SheetFile := SheetXml2}.

% @doc Re-assembles the XLSX zip archive exporting all XML elements in binaries.
-spec render(Filename :: file:name(), Xlsx :: map()) ->
    {ok, {FileName :: file:name(), binary()}}
    | {error, Reason :: term()}.
render(Filename, Xlsx) ->
    ZipBinaries = #{K => export_xml(V) || K := V <- Xlsx},
    zip:create(Filename, maps:to_list(ZipBinaries), [memory]).

%- Internal --------------------------------------------------------------------

read_worksheet(Row, Column, #xmlElement{name = worksheet, content = Content}) ->
    case get_xml_elem_by_name(sheetData, Content) of
        undefined -> {error, no_sheetdata};
        SheetData -> read_row_in_worksheet(Row, Column, SheetData)
    end.

read_row_in_worksheet(Row, Column, #xmlElement{content = Rows}) ->
    case get_xml_elem_by_attr(r, Row, Rows) of
        undefined -> {error, {row_not_found, Row}};
        RowElement  -> read_column_in_row(Column, RowElement)
    end.

read_column_in_row(Column, #xmlElement{content = Columns}) ->
    case get_xml_elem_by_attr(r, Column, Columns) of
        undefined -> {error, {column_not_found, Column}};
        CellElement  -> read_value_in_cell(CellElement)
    end.

read_value_in_cell(#xmlElement{content = CellContent}) ->
    case get_xml_elem_by_name(v, CellContent) of
        undefined -> empty;
        #xmlElement{name = v,
                    content = [#xmlText{type = text, value = V}]} -> V
    end.

get_xml_elem_by_name(Name, Content) ->
    case [E || #xmlElement{name = N} = E <- Content, N == Name] of
        [] -> undefined;
        [E|_]  -> E
    end.

get_xml_elem_by_attr(AttrName, AttrValue, Content) ->
    case [E || E <- Content, xml_elem_has_attr(AttrName, AttrValue, E)] of
        [] -> undefined;
        [E|_]  -> E
    end.

write_worksheet(Row, Column, Value,
                    #xmlElement{name = worksheet, content = Content} = WS) ->
    Fun = fun
        (#xmlElement{name = 'sheetData'} = D) ->
            write_worksheet_data(Row, Column, Value, D);
        (E) -> E
    end,
    WS#xmlElement{content = lists:map(Fun, Content)}.

write_worksheet_data(RowNumber, Column, Value,
                    #xmlElement{name = sheetData, content = Rows} = SD) ->
    Fun = fun(Row) ->
        case xml_elem_has_attr('r', RowNumber, Row) of
            true -> write_column_in_row(Column, Value, Row);
            false -> Row
        end
    end,
    SD#xmlElement{content = lists:map(Fun, Rows)}.

write_column_in_row(ColumID, Value,
                   #xmlElement{name = row, content = Columns} = Row) ->
    Fun = fun(Column) ->
        case xml_elem_has_attr('r', ColumID, Column) of
            true -> write_value_in_cell(Value, Column);
            false -> Column
        end
    end,
    Row#xmlElement{content = lists:map(Fun, Columns)}.

write_value_in_cell(Value, #xmlElement{content = Content} = Cell) ->
    OtherContent = [A || #xmlElement{name = N} = A <- Content, N /= v],
    XmlValue = #xmlElement{
        name = v,
        content = [
            #xmlText{type = text, pos = 1, value = value_to_string(Value)}
    ]},
    Cell1 = set_cell_type(Value, Cell),
    Cell1#xmlElement{content = OtherContent ++ [XmlValue]}.

set_cell_type(Value, #xmlElement{attributes = Attrs} = Cell)
when is_binary(Value) or is_list(Value) ->
    OtherAttr = [A || #xmlAttribute{name = N} = A <- Attrs, N /= t],
     Cell#xmlElement{
            attributes = [
                #xmlAttribute{name = t, value = "str"} | OtherAttr]
    };
set_cell_type(_, Cell) -> % type is not set if the value is a number
    Cell.

value_to_string(V) when is_list(V) ; is_binary(V) ->
    V;
value_to_string(V) when is_integer(V) ->
    integer_to_list(V);
value_to_string(V) when is_float(V) ->
    float_to_list(V).

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
    unicode:characters_to_binary(xmerl:export_simple([E], xmerl_xml)).

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
