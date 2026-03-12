type coordinate = Cord of int * int
[@@deriving show {with_path = false}]
type wall  = Wall of coordinate * coordinate
[@@deriving show {with_path = false}]

(*Ax+By=C*)
type line = {equation : int list; coord1: coordinate; coord2:coordinate;}

let save_line w = 
  match w with
  | Wall (Cord (x1,y1), Cord (x2,y2)) ->
      let a = y2 - y1 in
      let b = x1 - x2 in
      let c = a*x1 + b*y1 in
      {
        equation = [a;b;c];
        coord1 = Cord(x1,y1);
        coord2 = Cord(x2,y2)
      }


let rec map f list = match list with
|x::xs -> (f x)::( map f xs)
|[] ->[]
|_->[]


let range_overlap list =
  match list with
  | [Cord(xa1,ya1); Cord(xa2,ya2);
     Cord(xb1,yb1); Cord(xb2,yb2)] ->

       (
        min (max xa1 xa2) (max xb1 xb2) >
        max (min xa1 xa2) (min xb1 xb2)
        ||
        min (max ya1 ya2) (max yb1 yb2) >
        max (min ya1 ya2) (min yb1 yb2)
      )

  | _ -> false


let same_slope line1 line2= match line1,line2 with
|{equation = [a1;b1;c1]; coord1 = x1; coord2 = y1},{equation = [a2;b2;c2]; coord1 = x2; coord2 =y2 }-> if (a1*c2 = c1*a2 && a1*b2 = b1*a2 && c1*b2=b1*c2) then (if (range_overlap [x1;y1;x2;y2]) then false else true) else true
|_,_ -> false

let diffrent_slope  line1 line2= match line1,line2 with
|{equation = [a1;b1;c1]; coord1 = Cord (xa1, ya1); coord2 = Cord (xa2, ya2)},{equation = [a2;b2;c2]; coord1 = Cord (xb1, yb1); coord2 =Cord (xb2, yb2)}-> let d =float_of_int ((a1*b2)-(a2*b1)) in let x = (float_of_int ((c1*b2)-(b1*c2)))/.d in let y= (float_of_int ((c1*a2)-(c2*a1)))/.d in (if ((x <float_of_int (max xa1 xa2)) && x> float_of_int (min xa1 xa2))&&(( x>float_of_int (min xb1 xb2)&& x<float_of_int (max xb1 xb2)) ||(y<(float_of_int (max ya1 ya2))&& y>(float_of_int ( min ya1 ya2)))&& ((y<(float_of_int (max yb1 yb2))&& y>(float_of_int ( min yb1 yb2))))) then false else true)
|_,_ -> false

(*check crossing*)
let compare_line f line1 line2 =match line1,line2 with
|{equation = [a1;b1;c1]; coord1 = x1; coord2 = y1},{equation = [a2;b2;c2]; coord1 = x2; coord2 =y2 }-> if ((a1 *b2)-(b1 * a2)) = 0 then (same_slope {equation = [a1;b1;c1]; coord1 = x1; coord2 = y1} {equation = [a2;b2;c2]; coord1 = x2; coord2 =y2}) else (diffrent_slope {equation = [a1;b1;c1]; coord1 = x1; coord2 = y1} {equation = [a2;b2;c2]; coord1 = x2; coord2 =y2 })
|_-> false


let rec iterate_line f list = let rec helper f line list1 =match list1 with
|x::xs -> (compare_line f line x) :: (helper f line xs)
|[]->[]
in
 match list with
|{equation = [a1;b1;c1]; coord1 = x1; coord2 = y1}::xs -> (helper f {equation = [a1;b1;c1]; coord1 = x1; coord2 = y1} xs ) @ (iterate_line f xs)
|[] ->[]
|_->[]


let rec fold_left f acc list = match list with
| [] -> acc
| h :: t -> fold_left f (f acc h) t


let is_optimal walls =
  fold_left (&&) true (iterate_line compare_line (map save_line walls))


let point_line_cross point line = match point,line with
|Cord(a,b),Wall(Cord(a1,b1),Cord(a2,b2))-> if ((b-b1)*(a2-a1)-(a-a1)*(b2-b1) = 0) then (if ((min a1 a2) <= a && a<= (max a1 a2))&&((min b1 b2)<=b && b<=(max b1 b2)) then true else false) else false
|_->true

let detect_start_point path board= match path,board with
|x::xs,y::ys-> point_line_cross x y 
|x,y ->false
|[],[]->false 


let detect_end_point path board= match board with
|ya::y::ys->  let rec help path = match path with
|Cord(a,b)::[]->Cord(a,b)
|x::xs-> help xs
|[]->Cord(-9999,-9999)  in 
point_line_cross (help path) y
|_->false 

let rec calculate_velocity list =match list with
|Cord(a,b)::[]-> []
|Cord(a,b)::Cord(c,d)::xs -> Cord(c-a,d-b)::(calculate_velocity (Cord(c,d)::xs))
|[]->[]

let rec valid_velocity list= match list with
|Cord(a,b)::[]-> if (-1<= a && a<= 1)&&(-1<=b &&b<=1) then true else false
|Cord(a,b)::Cord(c,d)::xs -> if (a-c<=1 &&a-c>=(-1))&&(b-d<=1 &&b-d>=(-1)) then (valid_velocity (Cord(c,d)::xs)) else false
|[]->true

let rec merge_to_wall_list path board = match path with
|Cord(a,b)::Cord(c,d)::xs-> Wall(Cord(a,b),Cord(c,d))::(merge_to_wall_list (Cord(c,d)::xs) board)
|_-> board

let is_valid_path path board = (detect_start_point path board)&&(detect_end_point path board)&&(valid_velocity (calculate_velocity path))&& (is_optimal (merge_to_wall_list path board))