# Xlerl

Standalone minimal XLSX editing library in erlang.

Xlerl allows to quickly edit any XLSX file. Data is unpacked with zip and internal XML files are parsed and written with xmerl.

## Features
This library was built with the simple requirement of editing existing xlsx templates by reading and writing in arbitrary cells on arbitrary sheets.

- Open the xlsx zip file parsing *ALL* xml elements with xmerl.
- Read single cells
- Write single cells
- Export the edited XML back into a zip file

## Example

Parse, read, write and then export.

```erlang
    Filename = "spreadsheets.xlsx",
    {ok, Binary} = file:read_file(Filename),
    % Parse
    {ok, Xlsx} = xlerl:parse({Filename, Binary}),
    % You get a map of the whole zip, like:
    % #{InternalFilenamePath := XmerlParsedContent} = Xlsx,
    % Read
    V = case xlerl:read("MySheet", "A", "1", Xlsx) of
        empty -> 0.0;
        {error,E} -> error(E);
        V -> V
    end,
    % Write
    NewXlsx = xlerl:write("MySheet", "A", "1", V + 1.0, Xlsx),
    % Export and recreate the zip
    {ok, {Filename, NewBinary}} = xlerl:render(Filename, NewXlsx),
    file:write_file(Filename, NewBinary),
```
