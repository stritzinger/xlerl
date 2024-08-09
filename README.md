# Xlerl

<p >
  <a href="https://github.com/stritzinger/xlerl/actions/workflows/ci.yml?query=branch%3Amain">
    <img alt="continous integration" src="https://img.shields.io/github/actions/workflow/status/stritzinger/xlerl/ci.yml?label=build&style=flat-square&branch=main"/>
  </a>
  <a href="https://hex.pm/packages/xlerl">
    <img alt="hex.pm version" src="https://img.shields.io/hexpm/v/xlerl?style=flat-square"/>
  </a>
  <a href="LICENSE.md">
    <img alt="hex.pm license" src="https://img.shields.io/hexpm/l/xlerl?style=flat-square"/>
  </a>
  <a href="https://github.com/stritzinger/xlerl/blob/main/.github/workflows/ci.yml#L18">
    <img alt="erlang versions" src="https://img.shields.io/badge/erlang-26+-blue.svg?style=flat-square"/>
  </a>
</p>

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
