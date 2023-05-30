port module Spreadsheet exposing (..)

import Browser
import Browser.Dom
import Browser.Events
import Dict exposing (Dict)
import Html exposing (Html, div)
import Html.Attributes as HA
import Html.Events as HE
import Json.Decode as JD
import Json.Encode as JE
import Parser as P exposing ((|.), (|=), Parser)
import Task
import Tuple exposing (first, second, pair)

fi = String.fromInt
ff = String.fromFloat
tf = toFloat

port receive_spreadsheet : (JE.Value -> msg) -> Sub msg
port receive_attributes : (JE.Value -> msg) -> Sub msg
port send_spreadsheet : JE.Value -> Cmd msg

spreadsheet_updated : Model -> (Model, Cmd Msg)
spreadsheet_updated model = (model, encode_spreadsheet model.spreadsheet |> send_spreadsheet)

-- coords are (column, row)
type alias Coords = (Int,Int)

-- a range is (topleft, bottomright)
type alias Range = (Coords, Coords)

range_to_string ((x1,y1),(x2,y2)) = (fi x1)++"-"++(fi y1)++"--"++(fi x2)++"-"++(fi y2)

type alias Spreadsheet =
    { content : List (Range, CellContent)
    , range : Range
    }

type alias CellContent = 
    { type_ : CellType
    , content : String
    , style : CellStyle
    , disabled : Bool
    }

type CellType
    = NumberCell
    | StringCell

type alias CellStyle = List (String,String)

main = Browser.element
    { init = init
    , update = update
    , subscriptions = subscriptions
    , view = view
    }

type alias Model = 
    { spreadsheet: Spreadsheet
    , pressing : Maybe Coords
    , selection : Maybe Range
    , hover_cell : Maybe (Range, Coords)
    , disabled : Bool
    }

type alias SpreadsheetContext = 
    { selection : Maybe Range
    , disabled : Bool
    }


type Msg
    = Noop
    | InvalidPortMessage JD.Value
    | SetDisabled Bool
    | ReceiveSpreadsheet JD.Value
    | AddRowBefore Int
    | AddColumnBefore Int
    | DeleteRow Int
    | DeleteColumn Int
    | SetValue Coords String
    | MouseMoveCell Range Browser.Dom.Element
    | HoverLeaveCell Range
    | PressCell
    | MouseUp
    | ClickOutside
    | MoveSelection (Int,Int)
    | SetSelection Range

init : () -> (Model, Cmd Msg)
init _ = (init_model, Cmd.none)

empty_cell = { type_ = StringCell, content = "", style = [], disabled = False }
plain_cell range content = (range, { empty_cell | content = content })

set_cell_type type_ cell = { cell | type_ = type_ }

set_cell_content content cell = { cell | content = content }

set_cell_disabled disabled cell = { cell | disabled = disabled }

set_cell_style style cell = { cell | style = style }

empty_spreadsheet = { range = ((0,0),(0,0)), content = [] }

init_spreadsheet : Spreadsheet
init_spreadsheet = empty_spreadsheet

init_model = 
    { spreadsheet = init_spreadsheet
    , pressing = Nothing
    , selection = Nothing
    , hover_cell = Nothing
    , disabled = False
    }

nocmd model = (model, Cmd.none)

update : Msg -> Model -> (Model, Cmd Msg)
update msg model = case msg of
    Noop -> nocmd model
    InvalidPortMessage _ -> model |> nocmd
    SetDisabled disabled -> { model | disabled = disabled } |> nocmd
    ReceiveSpreadsheet v -> 
        case JD.decodeValue decode_spreadsheet v of
            Ok spreadsheet -> { init_model | spreadsheet = spreadsheet, disabled = model.disabled } |> spreadsheet_updated
            Err err -> model |> nocmd
    AddRowBefore row -> update_spreadsheet (add_row_before row) model |> spreadsheet_updated
    AddColumnBefore column -> update_spreadsheet (add_column_before column) model |> spreadsheet_updated
    DeleteRow row -> update_spreadsheet (delete_row row) model |> spreadsheet_updated
    DeleteColumn column -> update_spreadsheet (delete_column column) model |> spreadsheet_updated
    SetValue pos value -> update_spreadsheet (set_value pos value) model |> spreadsheet_updated
    MouseMoveCell range e -> 
        let
            ((x1,y1),(x2,y2)) = range
            w = 1 + (tf (x2-x1))
            h = 1 + (tf (y2-y1))
            x = e.viewport.x
            y = e.viewport.y
            (dx, dy) = (x - e.element.x, y - e.element.y)
            (mx,my) = (floor (dx*w/e.element.width), floor (dy*h/e.element.height))
        in
            { model | hover_cell = Just (range, (mx + x1, my + y1)) } |> update_selection |> nocmd

    HoverLeaveCell range -> (if Maybe.map first model.hover_cell == Just range then { model | hover_cell = Nothing} else model) |> nocmd
    PressCell -> case model.hover_cell of
        Just (_,coords) -> { model | pressing = Just coords, selection = Just (coords,coords) } |> nocmd
        Nothing -> model |> nocmd
    MouseUp -> { model | pressing = Nothing } |> nocmd
    ClickOutside -> (if model.hover_cell == Nothing then { model | selection = Nothing } else model) |> nocmd
    MoveSelection (dx,dy) -> case model.selection of
        Just selection -> 
            let
                ((x1,y1),(x2,y2)) = case selected_cells model |> List.head of
                    Just (r,_) -> r
                    Nothing -> selection
                ((rx1,ry1),(rx2,ry2)) = model.spreadsheet.range
                (x,y) = (if dx<=0 then x1 else x2, if dy<=0 then y1 else y2)
                pos = (clamp rx1 rx2 <| x+dx, clamp ry1 ry2 <| y+dy)
            in
                { model | selection = Just (pos,pos) } |> nocmd

        Nothing -> model |> nocmd

    SetSelection range -> { model | selection = Just range } |> nocmd

update_selection model = case (model.pressing, model.hover_cell) of
    (Just (x1,y1), Just (_,(x2,y2))) -> { model | selection = Just ((min x1 x2, min y1 y2), (max x1 x2, max y1 y2)) }
    _ -> model

update_spreadsheet : (Spreadsheet -> Spreadsheet) -> Model -> Model
update_spreadsheet fn model = { model | spreadsheet = fn model.spreadsheet }

subscriptions model = Sub.batch
    [ Browser.Events.onMouseUp (JD.succeed MouseUp)
    , Browser.Events.onClick (JD.succeed ClickOutside)
    , receive_spreadsheet ReceiveSpreadsheet
    , receive_attributes decode_received_attribute
    ]

decode_received_attribute : JE.Value -> Msg
decode_received_attribute v = 
    v
    |>
    JD.decodeValue (
        JD.oneOf
        [ JD.field "disabled" JD.bool |> JD.map SetDisabled
        ]
    )
    |> Result.withDefault (InvalidPortMessage v)

selected_cells : Model -> List (Range, CellContent)
selected_cells model = case model.selection of
    Just range -> List.filter (first >> ranges_overlap range) model.spreadsheet.content
    Nothing -> []

view model = 
    let
        context = { selection = model.selection, disabled = model.disabled }
    in
        div
            []
            [ view_spreadsheet context model.spreadsheet
            ]
            {-
            , Html.pre [] [Html.text <| JE.encode 2 (encode_spreadsheet model.spreadsheet)]
            , Html.p [] [Html.text "hovering ", Html.text <| Debug.toString <| model.hover_cell]
            , Html.p [] [Html.text "pressing ", Html.text <| Debug.toString <| model.pressing]
            , Html.p [] [Html.text "selection ", Html.text <| Debug.toString <| model.selection]
            , Html.p [] [Html.text "selected cells", Html.text <| Debug.toString <| selected_cells model]
            , Html.p [] [Html.text "content", Html.text <| Debug.toString <| List.map (\(range,c) -> (range,c.content)) model.spreadsheet.content]
            , Html.input [ HE.onInput SetInput, HA.value model.input ] []
            , Html.p [] [Html.text <| Debug.toString <| JD.decodeString decode_cell_style model.input ]
            ]
            -}

range_width : Range -> Int
range_width ((x1,y1),(x2,y2)) = (abs (x2-x1))+1

range_height : Range -> Int
range_height ((x1,y1),(x2,y2)) = (abs (y2-y1))+1

format_cell_value : CellContent -> String
format_cell_value cell = case cell.type_ of
    NumberCell -> String.toFloat cell.content |> Maybe.map ff |> Maybe.withDefault "ERR"
    StringCell -> cell.content

view_spreadsheet : SpreadsheetContext -> Spreadsheet -> Html Msg
view_spreadsheet context spreadsheet =
    let
        ((x1,y1),(x2,y2)) = spreadsheet.range
    in
        Html.table
            []
            (List.map (\y -> Html.tr [] (List.filterMap (view_cell_corner context spreadsheet y) (List.range x1 x2))) (List.range y1 y2))

range_contains : Coords -> Range -> Bool
range_contains (x,y) ((x1,y1),(x2,y2)) = x>=x1 && x<=x2 && y>=y1 && y<=y2

ranges_overlap : Range -> Range -> Bool
ranges_overlap ((ax1,ay1),(ax2,ay2)) ((bx1,by1),(bx2,by2)) = ax1 <= bx2 && ax2 >= bx1 && ay1 <= by2 && ay2 >= by1

-- does range a entirely contain range b?
range_surrounds : Range -> Range -> Bool
range_surrounds ((ax1,ay1),(ax2,ay2)) ((bx1,by1),(bx2,by2)) = ax1 <= bx1 && ax2 >= bx2 && ay1 <= by1 && ay2 >= by2

get_containing_cell : List (Range,CellContent) -> Coords -> Maybe (Range,CellContent)
get_containing_cell content pos =
    content
    |> List.filter (first >> range_contains pos)
    |> List.head

view_cell_corner : SpreadsheetContext -> Spreadsheet -> Int -> Int -> Maybe (Html Msg)
view_cell_corner context spreadsheet y x = case get_containing_cell spreadsheet.content (x,y) of
    Just (((x1,y1), (x2,y2)) as range, content) -> if x1==x && y1==y then Just (view_cell context range content) else Nothing
    Nothing -> Just <| view_cell context ((x,y),(x,y)) empty_cell

view_cell : SpreadsheetContext -> Range -> CellContent -> Html Msg
view_cell context range cell =
    let
        id = "cell-"++(range_to_string range)

        in_selection = 
               context.selection
            |> Maybe.andThen (\s -> if ranges_overlap range s then Just True else Nothing)
            |> ((/=) Nothing)

        input_active = Maybe.map (range_surrounds range) context.selection == Just True

        num_lines = cell.content |> String.split("\n") |> List.length

        events = 
            [ HE.onFocus (SetSelection range)
            , HA.class "input-value"
            , HE.on "cellup" (JD.succeed <| MoveSelection (0,-1))
            , HE.on "celldown" (JD.succeed <| MoveSelection (0,1))
            , HE.on "cellleft" (JD.succeed <| MoveSelection (-1,0))
            , HE.on "cellright" (JD.succeed <| MoveSelection (1,0))
            ]

        disabled = cell.disabled || context.disabled
    in
        Html.td
            ( ((List.map (\(k,v) -> HA.style k v)) cell.style)
              ++
              (if input_active then [ HA.attribute "data-input-active" "true" ] else [])
              ++
              [ HA.colspan <| range_width range
              , HA.rowspan <| range_height range
              , HA.id id
              , HA.attribute "data-annotated-mousemove" "true"
              , HE.on "annotatedmousemove" (JD.map (MouseMoveCell range) decode_annotated_mousemove)
              , HE.onMouseOut <| HoverLeaveCell range
              , HE.onMouseDown <| PressCell
              , HA.classList
                [ ("in-selection", in_selection)
                , ("input-active", input_active)
                ]
              ]
            )

            ( if disabled then
                [ Html.div
                    (   [ HA.attribute "tabindex" "0"
                        , HA.class "input-value disabled"
                        ]
                        ++
                        events
                    )
                    [ Html.text cell.content ]
                ]
              else
                [ Html.textarea
                    (
                        [ HA.value cell.content
                        , HE.onInput (SetValue (first range))
                        , HA.style "height" <| (fi (num_lines+1))++"em"
                        ] 
                        ++
                        events
                    )
                    [] 
                ]
            )

add_row_before : Int -> Spreadsheet -> Spreadsheet
add_row_before row spreadsheet = 
    let
        content = 
               spreadsheet.content
            |> List.map (\(((x1,y1),(x2,y2)),cell) ->
                (((x1,y1+(if y1>=row then 1 else 0)), (x2,y2+(if y1>=row then 1 else 0)+(if row>y1 && row<=y2 then 1 else 0))), cell)
               )

        range = 
            let
                ((x1,y1),(x2,y2)) = spreadsheet.range
            in
                ((x1,y1),(x2,y2+1))
    in
        { range = range, content = content }

delete_row : Int -> Spreadsheet -> Spreadsheet
delete_row row spreadsheet =
    let
        fix rr = rr - (if rr>=row then 1 else 0)

        content =
               spreadsheet.content
            |> List.filter (\(((x1,y1),(x2,y2)), cell) -> not (y1==row && y2==row))
            |> List.map (\(((x1,y1),(x2,y2)), cell) -> (((x1,y1),(x2,fix y2)), cell))

        range =
            let
                ((x1,y1),(x2,y2)) = spreadsheet.range
            in
                ((x1,max 0 (fix y1)), (x2, fix y2))

    in
        { range = range, content = content }

add_column_before : Int -> Spreadsheet -> Spreadsheet
add_column_before col spreadsheet =
    let
        content =
               spreadsheet.content
            |> List.map (\(((x1,y1),(x2,y2)),cell) ->
                (((x1+(if x1>=col then 1 else 0), y1), (x2+(if x1>=col then 1 else 0)+(if col>x1 && col<=x2 then 1 else 0),y2)),cell)
               )

        range =
            let
                ((x1,y1),(x2,y2)) = spreadsheet.range
            in
                ((x1,y1),(x2+1,y2))
    in
        { range = range, content = content }

delete_column : Int -> Spreadsheet -> Spreadsheet
delete_column column spreadsheet =
    let
        fix cc = cc - (if cc >= column then 1 else 0)

        content =
               spreadsheet.content
            |> List.filter (\(((x1,y1),(x2,y2)), cell) -> not (x1==column && x2==column))
            |> List.map (\(((x1,y1),(x2,y2)), cell) ->
                (((if x1==column then column else fix x1, y1), (fix x2, y2)), cell)
               )

        range =
            let
                ((x1,y1),(x2,y2)) = spreadsheet.range
            in
                ((max 0 (fix x1), y1), (max 0 (fix x2), y2))

    in
        { range = range, content = content }

set_value : Coords -> String -> Spreadsheet -> Spreadsheet
set_value pos cell_content spreadsheet =
    let
        existing_cell = get_containing_cell spreadsheet.content pos

        content = case existing_cell of
            Just ((r,cell) as ec) -> 
                List.map 
                    (\c ->
                        if c == ec then
                            (r, { cell | content = cell_content })
                        else
                            c
                    )
                    spreadsheet.content

            Nothing -> ((pos,pos), { empty_cell | content = cell_content })::spreadsheet.content

    in
        { spreadsheet | content = content }

alpha_to_num : String -> Int
alpha_to_num = String.toList >> List.foldl (\c -> \n -> n * 26 + (Char.toCode c) + 1 - (Char.toCode 'A')) 0

num_to_alpha : Int -> String
num_to_alpha n = 
    let
        m = remainderBy 26 n
        k = n // 26
        c = String.fromChar <| Char.fromCode (m + (Char.toCode 'A'))
    in
        if k > 0 then (num_to_alpha (k-1))++c else c

parse_coords : Parser Coords
parse_coords =
    P.succeed (\c -> \r -> (c-1,r-1))
     |= (P.mapChompedString 
          (\s -> \_ -> alpha_to_num s)
          ( P.succeed ()
            |. P.chompIf Char.isUpper
            |. P.chompWhile Char.isUpper
          )
        )

     |= P.int

parse_range : Parser Range
parse_range =
    P.succeed pair
    |= parse_coords
    |. P.symbol ":"
    |= parse_coords

decode_spreadsheet : JD.Decoder Spreadsheet
decode_spreadsheet =
    JD.map3 
        (\content -> \range -> \merges ->
            let
                kept_content = 
                       content
                    |> List.filterMap 
                        (\(pos,cell) -> 
                            let
                                overlaps = List.filterMap (\r -> if range_contains pos r then Just r else Nothing) merges
                                crange = case overlaps of
                                    r::[] -> if first r == pos then Just r else Nothing
                                    [] -> Just (pos,pos)
                                    _ -> Nothing
                            in 
                                Maybe.map (\r -> (r,cell)) crange
                        )
            in
                { content = kept_content, range = range }
        )
        decode_spreadsheet_content
        (JD.field "!ref" decode_range)
        (JD.oneOf 
            [ JD.field "!merges" (JD.list decode_range)
            , JD.succeed []
            ]
        )

decode_range : JD.Decoder Range
decode_range = 
    JD.oneOf
        [ JD.string
            |> JD.andThen (\s ->
                case P.run parse_range s of
                    Ok r -> JD.succeed r
                    Err _ -> JD.fail <| "invalid range: "++s
                )

        , JD.map2 pair
            (JD.field "s" decode_coords_object)
            (JD.field "e" decode_coords_object)
        ]

decode_coords : JD.Decoder Coords
decode_coords = 
    JD.string
    |> JD.andThen (\s ->
        case P.run parse_coords s of
            Ok c -> JD.succeed c
            Err _ -> JD.fail <| "invalid coords: "++s
        )

decode_coords_object : JD.Decoder Coords
decode_coords_object =
    JD.map2 pair
        (JD.field "c" JD.int)
        (JD.field "r" JD.int)

decode_spreadsheet_content : JD.Decoder (List (Coords, CellContent))
decode_spreadsheet_content = 
    JD.dict JD.value
    |> JD.map Dict.toList
    |> JD.map (List.filterMap (\(k,v) -> Maybe.map (\r -> (r,v)) (P.run parse_coords k |> Result.toMaybe)))
    |> JD.map (List.filterMap 
        (\(k,v) -> case JD.decodeValue decode_cell v of
            Ok c -> Just (k,c)
            Err _ -> Nothing
        )
       )

decode_any_of : List (JD.Decoder a) -> JD.Decoder (List a)
decode_any_of decoders = List.foldr
    (\d -> JD.andThen (\rest ->
           JD.maybe d 
        |> JD.map (\mv -> case mv of
            Just v -> v::rest
            Nothing -> rest
           )
    ))
    (JD.succeed [])
    decoders

decode_cell_style : JD.Decoder CellStyle
decode_cell_style =
    (decode_any_of >> JD.map (List.concatMap identity))
    [ JD.field "border" decode_border_style
    , JD.field "font" decode_font_style
    , JD.field "fill" decode_fill_style
    , JD.field "alignment" decode_alignment_style
    ]

decode_border_style : JD.Decoder CellStyle
decode_border_style =
    (decode_any_of >> JD.map (List.concatMap identity))
        (List.map 
            (\side -> JD.field side (decode_border_side side))
            ["top","bottom","left","right","diagonal"]
        )

decode_border_side : String -> JD.Decoder CellStyle
decode_border_side side =
       JD.field "style" decode_border_size
    |> JD.andThen (\size -> 
        JD.map (\c -> [size,c])
        (JD.oneOf
            [ JD.field "color" decode_color |> JD.map (\c -> ("color",c))
            , JD.succeed ("color","black")
            ]
        )
        )
    |> JD.map (List.map (\(k,v) -> ("border-"++side++"-"++k, v)))

decode_border_size =
    JD.string
    |> JD.andThen (\s -> case s of
        "thin" -> JD.succeed "0.1rem"
        "medium" -> JD.succeed "0.2rem"
        "thick" -> JD.succeed "0.3rem"
        _ -> JD.fail <| "Invalid border size: "++s
       )
    |> JD.map (\s -> ("width",s))


decode_color : JD.Decoder String
decode_color = 
    JD.oneOf
        [ JD.field "rgb" JD.string |> JD.map (\c -> 
            if String.length c == 8 then 
                "#"++(String.slice 2 8 c)++(String.slice 0 2 c)
            else
                "#"++c
          )
        , JD.field "css" JD.string
        ]

decode_bool_int_field name v = 
    JD.field name 
    (JD.oneOf
        [ JD.int |> JD.map ((==) 1)
        , JD.bool
        ]
    )
    |> JD.andThen (\b -> if b then JD.succeed v else JD.fail ("not "++name))

decode_font_style =
    decode_any_of
        [ JD.field "name" JD.string |> JD.map (\f -> ("font-family",f))
        , JD.field "sz" JD.float |> JD.map (\sz -> ("font-size", (ff (sz/11)++"em")))
        , decode_bool_int_field "italic" ("font-style", "italic")
        , decode_bool_int_field "bold" ("font-weight","bold")
        , decode_bool_int_field "underline" ("text-decoration", "underline")
        , JD.field "color" decode_color |> JD.map (\c -> ("color",c))
        ]

decode_fill_style =
    decode_any_of
        [ JD.field "fgColor" decode_color |> JD.map (\c -> ("background-color",c))
        ]

decode_alignment_style =
    decode_any_of
        [ JD.field "horizontal" JD.string |> JD.map (\c -> ("text-align",c))
        , JD.field "vertical" JD.string |> JD.map (\c -> ("vertical-align",c))
        ]

decode_cell : JD.Decoder CellContent
decode_cell =
    ( JD.field "t" JD.string
      |> JD.andThen
        (\t -> case t of
            "n" -> JD.succeed (NumberCell, JD.map ff JD.float)
            "s" -> JD.succeed (StringCell, JD.string)
            "z" -> JD.succeed (StringCell, JD.succeed "")
            _ -> JD.fail <| "Invalid cell type: "++t
        )
      |> JD.andThen
        (\(t,dv) -> JD.map (\v -> empty_cell |> set_cell_type t |> set_cell_content v) (JD.oneOf [JD.field "v" dv, JD.succeed ""]))
      |> JD.map2
        set_cell_disabled
        (JD.oneOf [JD.field "disabled" JD.bool, JD.succeed False])
      |> JD.map2
        set_cell_style
        (JD.oneOf [JD.field "style" decode_cell_style, JD.succeed []])
    )
        

decode_annotated_mousemove : JD.Decoder Browser.Dom.Element
decode_annotated_mousemove =
    JD.field "detail" <|
        JD.map3
            (\x -> \y -> \element -> { scene = { width = 0, height = 0 }, viewport = { x = x, y = y, width = 0, height = 0 }, element = element })
            (JD.field "x" JD.float)
            (JD.field "y" JD.float)
            (JD.field "box" <| JD.map4 (\x -> \y -> \width -> \height -> { x = x, y = y, width = width, height = height })
                (JD.field "x" JD.float)
                (JD.field "y" JD.float)
                (JD.field "width" JD.float)
                (JD.field "height" JD.float)
            )

debug = 
    JD.field "detail" (JD.dict (JD.value))
    |> JD.andThen (\d ->
        let
            element = { x = 0, y = 0, width = 0, height = 0 }
            e = { scene = { width = 0, height = 0 }, viewport = { x = 0, y = 0, width = 0, height = 0 }, element = element }
        in
            JD.succeed e
      )

encode_spreadsheet : Spreadsheet -> JE.Value
encode_spreadsheet spreadsheet = 
    JE.object
        ([ ("!ref", encode_range spreadsheet.range)
         , ("!merges", 
                spreadsheet.content
             |> List.map first
             |> List.filterMap (\(a,b) -> if a /= b then Just (a,b) else Nothing)
             |> JE.list encode_range
           )
         ] 
         ++
         (List.map encode_cell spreadsheet.content)
        )

encode_range : Range -> JE.Value
encode_range (a,b) = JE.string <| (coords_to_string a)++":"++(coords_to_string b)

coords_to_string : Coords -> String
coords_to_string (x,y) = (num_to_alpha (x))++(fi (y+1))

encode_cell: (Range,CellContent) -> (String, JE.Value)
encode_cell (range,cell) =
    ( coords_to_string <| first range
    , JE.object
        [ ("t", encode_cell_type cell.type_)
        , ("v", case cell.type_ of
                NumberCell -> case String.toFloat cell.content of
                    Just n -> JE.float n
                    Nothing -> JE.string "ERR"
                StringCell -> JE.string cell.content
          )
        ]
    )

encode_cell_type : CellType -> JE.Value
encode_cell_type type_ = JE.string (case type_ of
    NumberCell -> "n"
    StringCell -> "s"
    )
