namespace ExcelPlugin

open System
open System.Collections.Generic

type CellMap<'T> =
    {
        Header : string option
        Map : 'T -> Cell
    }

type Table =
    static member inline cell(value) = Cell.Text value
    static member inline cell(value) = Cell.Number value
    static member inline cell(value) = Cell.Boolean value
    static member inline cell(value) = Cell.Integer value
    static member inline cell(value) = Cell.DateTime value
    static member inline cell(value) = Cell.IntegerOption value
    static member inline cell(value) = Cell.TextOption value
    static member inline cell(value) = Cell.BooleanOption value
    static member inline cell(value) = Cell.NumberOption value
    static member inline cell(value: DateTimeOffset) = Cell.Text (value.ToString("o"))
    static member inline cell(value: DateTimeOffset option) =
        value
        |> Option.map (fun date -> date.ToString("o"))
        |> Table.cell

    static member inline create (headers: string list) (rows: Cell list list) =
        array2D [|
            [| for header in headers -> box header |]
            for row in rows do
                [| for value in row -> Cell.Raw value |]
        |]

    static member inline scalar cell =
        array2D [|
            [| Cell.Raw cell |]
        |]

    static member inline scalar (cell: string) =
        array2D [|
            [| box cell |]
        |]

    static member inline from (data: seq<'T>) (cells: CellMap<'T> list) =
        let headers = cells |> List.choose (fun cell -> cell.Header)
        if headers.Length = 0 then
            // there are no headers
            array2D [|
                for element in data -> [|
                    for cell in cells -> Cell.Raw (cell.Map element)
                |]
            |]

        else
            let availableHeaders = [
                for cell in cells -> defaultArg cell.Header ""
            ]

            array2D [|
                [| for header in availableHeaders -> box header |]
                for element in data do [|
                    for cell in cells -> Cell.Raw (cell.Map element)
                |]
            |]

[<AutoOpen>]
module Extenstions =
    let (=>) header cell = { cell with Header = Some header }

    type IEnumerable<'T> with
        member inline data.cell(map: 'T -> string) : CellMap<'T> =
            {
                Header = None
                Map = fun value -> Table.cell (map value)
            }

        member inline data.cell(map: 'T -> string option) : CellMap<'T> =
            {
                Header = None
                Map = fun value -> Table.cell (map value)
            }

        member inline data.cell(map: 'T -> bool) : CellMap<'T> =
            {
                Header = None
                Map = fun value -> Table.cell (map value)
            }
        member inline data.cell(map: 'T -> bool option) : CellMap<'T> =
            {
                Header = None
                Map = fun value -> Table.cell (map value)
            }
        member inline data.cell(map: 'T -> int) : CellMap<'T> =
            {
                Header = None
                Map = fun value -> Table.cell (map value)
            }
        member inline data.cell(map: 'T -> int option) : CellMap<'T> =
            {
                Header = None
                Map = fun value -> Table.cell (map value)
            }
        member inline data.cell(map: 'T -> double) : CellMap<'T> =
            {
                Header = None
                Map = fun value -> Table.cell (map value)
            }
        member inline data.cell(map: 'T -> double option) : CellMap<'T> =
            {
                Header = None
                Map = fun value -> Table.cell (map value)
            }
        member inline data.cell(map: 'T -> decimal) : CellMap<'T> =
            {
                Header = None
                Map = fun value -> Table.cell (float(map value))
            }
        member inline data.cell(map: 'T -> decimal option) : CellMap<'T> =
            {
                Header = None
                Map = fun value ->
                    match map value with
                    | Some value -> Table.cell (float value)
                    | None -> Table.cell(None: float option)
            }

        member inline data.cell(map: 'T -> DateTime) : CellMap<'T> =
            {
                Header = None
                Map = fun value -> Table.cell (map value)
            }

        member inline data.cell(map: 'T -> DateTimeOffset) : CellMap<'T> =
            {
                Header = None
                Map = fun value -> Table.cell (map value)
            }
        member inline data.cell(map: 'T -> DateTimeOffset option) : CellMap<'T> =
            {
                Header = None
                Map = fun value -> Table.cell (map value)
            }