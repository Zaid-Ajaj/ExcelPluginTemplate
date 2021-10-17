namespace ExcelPlugin

open System

[<RequireQualifiedAccess>]
type Cell =
    | Integer of int
    | Text of string
    | Number of float
    | Boolean of bool
    | DateTime of DateTime
    | IntegerOption of int option
    | TextOption of string option
    | NumberOption of float option
    | BooleanOption of bool option

    static member inline Raw = function
        | Integer value -> box value
        | Text value -> box value
        | Number value -> box value
        | Boolean value -> box value
        | DateTime value -> box value
        | IntegerOption value ->
            match value with
            | None -> box String.Empty
            | Some value -> box value
        | NumberOption value ->
            match value with
            | None -> box String.Empty
            | Some value -> box value
        | TextOption value ->
            match value with
            | None -> box String.Empty
            | Some value -> box value
        | BooleanOption value ->
            match value with
            | None -> box String.Empty
            | Some value -> box value