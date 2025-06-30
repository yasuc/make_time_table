open Eio.Std
open Base

module Date = Ptime

(* キャッシュ制御 *)
let update_needed xlsx pkl =
  match Sys.file_exists pkl with
  | `Yes ->
    let x_m = Unix.stat xlsx |> fun s -> s.st_mtime in
    let p_m = Unix.stat pkl |> fun s -> s.st_mtime in
    x_m > p_m
  | _ -> true

(* バイナリキャッシュの読み書き *)
let save_cache file data =
  Out_channel.with_file file ~binary:true ~f:(fun oc ->
    Bin_prot.Common.to_channel oc (data : (Yojson.Basic.t array list) list))

let load_cache file =
  In_channel.with_file file ~f:(fun ic ->
    Bin_prot.Common.from_channel ic : (Yojson.Basic.t array list) list)

(* Excel読み込み＆キャッシュ *)
let process_schedule xlsx pkl =
  if update_needed xlsx pkl then
    Flambda (blabla)
  else
    load_cache pkl

(* 出力 *)
let print_schedule all2d =
  printf "Subject,Start Date,All Day Event\n";
  all2d |> List.iter ~f:(fun rows ->
    let current_day = ref "" in
    rows |> List.iter ~f:(fun row ->
      match Array.get row 0 with
      | `String s ->
        (* assume date string *)
        current_day := s
      | _ -> ();
      for j = 2 to 6 do
        match Array.get row j with
        | `String s ->
          let subj =
            Regex.replace ~pattern:"※.*" ~with_:"" s
            |> Regex.replace ~pattern:"[ 　]+" ~with_:""
          in
          if subj <> "" then
            printf "%s,%s,TRUE\n" subj !current_day
        | _ -> ()
      done))

let () =
  let xlsx, pkl =
    match Array.to_list Sys.argv |> List.tl with
    | x::y::_ -> x, y
    | x::_    -> x, "schedule.bin"
    | _       -> "schedule.xlsx", "schedule.bin"
  in
  Eio_main.run @@ fun env ->
    let feed = SZXX.Feed.of_flow (Eio.Path.open_in (Eio.Stdenv.fs env / xlsx)) in
    let seq =
      SZXX.Xlsx.stream_rows_double_pass
        ~sw:(Switch.create ~label:"xls" env)
        (Eio.Path.open_in (Eio.Stdenv.fs env / xlsx))
        SZXX.Xlsx.yojson_cell_parser
    in
    (* 必要な形に加工しながらバッファ *)
    let all2d = Sequence.to_list seq |> List.group ... in
    save_cache pkl all2d;
    print_schedule all2d
