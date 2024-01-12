-module(xlerl).
% @doc API to parse, edit and write xlsx files with xmerl and zip

% API
-export([parse/1]).
-export([add_shared_string/2]).
-export([edit/5]).
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
    {UniqueCount, Xlsx#{SharedStringsFile => NewSharedStrings}}.

edit(SheetName, Column, Row, Value, Xlsx) ->
    SheetFile = lookup_sheet_file(SheetName, Xlsx),
    #{SheetFile := SheetXml} = Xlsx,
    SheetXml2 = edit_worksheet(Column, Row, Value, SheetXml),
    Xlsx#{SheetFile := SheetXml2}.

render(Filename, Xlsx) ->
    ZipBinaries = #{K => export_xml(V) || K := V <- Xlsx},
    zip:create(Filename, maps:to_list(ZipBinaries), [memory]).

%- Internal --------------------------------------------------------------------

edit_worksheet(Column, Row, Value,
                    #xmlElement{name = worksheet, content = Content} = WS) ->
    Fun = fun
        (#xmlElement{name = 'sheetData'} = D) ->
            edit_worksheet_data(Column, Row, Value, D);
        (E) -> E
    end,
    WS#xmlElement{content = lists:map(Fun, Content)}.

edit_worksheet_data(Column, RowNumber, Value,
                    #xmlElement{name = sheetData, content = Rows} = SD) ->
    Fun = fun(Row) ->
        case xml_elem_has_attr('r', RowNumber, Row) of
            true -> edit_column_in_row(Column ++ RowNumber, Value, Row);
            false -> Row
        end
    end,
    SD#xmlElement{content = lists:map(Fun, Rows)}.

edit_column_in_row(ColumID, Value,
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
