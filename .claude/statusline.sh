#!/bin/zsh
# Statusline 2 dong: (1) Context | Model | Effort, (2) Quota.
# Tuy chinh: sua cac hang mau va gradient ngay duoi day.
#   NO_COLOR (co gia tri) tat toan bo ma mau.
# Ky tu thanh: BAR_FULL (phan day), BAR_EMPTY (phan rong). Thay the: ▄ (U+2584) ▬ (U+25AC) ━ (U+2501) ▰ (U+25B0)
BAR_FULL="■"; BAR_EMPTY="■"
# Dem nhan tinh: nhan "Context:" duoc dem de "[" dong 1 trung cot "[" dong 2 ("Quota: 5-Hours [").
QLABEL="Quota: 5-Hours"; CLABEL="Context:"
lpad="${(l:$(( ${#QLABEL} - ${#CLABEL} )):: :)}"
input=$(cat)
setopt extendedglob
# Dam bao zsh dem ky tu UTF-8 (thanh) thay vi byte khi locale la C/POSIX.
if [[ "${LC_ALL:-${LC_CTYPE:-$LANG}}" != *[uU][tT][fF]-#8* ]]; then
  [[ -n "$LC_ALL" ]] && export LC_ALL=en_US.UTF-8 || export LC_CTYPE=en_US.UTF-8
fi

# ---------- MAU (ma SGR 16 mau; sua tai day) ----------
C_MODEL="1;34"       # xanh duong dam
C_NEUTRAL="37"       # nhan Model/Effort/Context/Quota/5-Hours/Week
C_SEP="90"           # dau "|" va chi tiet, xam mo
# Phan tram theo nguong: <70 xanh duong, 70-89 cam, >=90 do (khong dung xanh la). 3 bien the theo che do mau.
C_OK_TRUE="38;2;59;130;246"; C_WARN_TRUE="38;2;249;115;22"; C_CRIT_TRUE="31"
C_OK_256="38;5;69";          C_WARN_256="38;5;208";         C_CRIT_256="31"
C_OK_16="94";                C_WARN_16="33";                C_CRIT_16="31"   # 16 mau khong co cam: dung vang (thuong hien cam/nau)
C_EFFORT_LOW="90"; C_EFFORT_MEDIUM="94"; C_EFFORT_HIGH="33"
C_EFFORT_XHIGH="91"; C_EFFORT_MAX="31"; C_EFFORT_OTHER="0"
# ---------- GRADIENT THANH (10 o: o dau xanh duong -> o cuoi cam) ----------
# truecolor: 4 diem dung (R G B) xanh duong, tim, hong, cam; noi suy tuyen tinh deu qua 10 o. Them/bot diem dung tuy y (>= 2).
GRAD_STOPS=("59 130 246" "139 92 246" "236 72 153" "249 115 22")
# 256 mau: 10 chi so tang dan, xanh duong -> tim -> hong -> cam; khong dung dai xanh la, khong cyan. (RGB bang xterm)
#   26=0,95,215 xanh duong | 27=0,95,255 xanh duong sang | 63=95,95,255 xanh tim | 99=135,95,255 tim
#   135=175,95,255 tim | 171=215,95,255 tim hong | 177=215,135,255 tim nhat | 199=255,0,175 hong
#   203=255,95,95 san ho | 208=255,135,0 cam (trung C_WARN_256)
GRAD_256=(26 27 63 99 135 171 177 199 203 208)
EMPTY_TRUE="38;2;85;85;85"; EMPTY_256="38;5;239"; EMPTY_16="90"   # o rong, xam mo
# -----------------------------------------------------------

# Che do mau: none | true | 256 | basic
if [[ -n "${NO_COLOR-}" ]]; then cmode=none
elif [[ "${COLORTERM-}" == (#i)(truecolor|24bit) ]]; then cmode=true
elif [[ "${TERM-}" == *256color* ]]; then cmode=256
else cmode=basic; fi
esc=$'\e'
# paint: $1 = ma SGR, $2 = chu. Co reset sau cum; khong mau thi in chu tran.
paint() { if [[ $cmode == none || -z "$2" ]]; then print -rn -- "$2"; else print -rn -- "${esc}[${1}m${2}${esc}[0m"; fi }
sep=" $(paint $C_SEP '|') "
# Nhom mau cua thanh: grad[1..10] (o day) va c_empty
typeset -a grad
c_empty=""
case $cmode in
  true)
    for i in {1..10}; do
      ns=$(( ${#GRAD_STOPS} - 1 ))            # so doan
      pos=$(( (i-1) * ns * 1000 / 9 ))        # 0..ns*1000
      sg=$(( pos / 1000 )); (( sg >= ns )) && sg=$(( ns - 1 ))
      u=$(( pos - sg * 1000 ))                # 0..1000 trong doan
      a=(${=GRAD_STOPS[sg+1]}); b=(${=GRAD_STOPS[sg+2]})
      r=$(( a[1] + (b[1]-a[1]) * u / 1000 )); g=$(( a[2] + (b[2]-a[2]) * u / 1000 )); bl=$(( a[3] + (b[3]-a[3]) * u / 1000 ))
      grad[i]="38;2;$r;$g;$bl"
    done
    c_empty=$EMPTY_TRUE; C_OK=$C_OK_TRUE; C_WARN=$C_WARN_TRUE; C_CRIT=$C_CRIT_TRUE;;
  256)
    for i in {1..10}; do grad[i]="38;5;${GRAD_256[i]}"; done
    c_empty=$EMPTY_256; C_OK=$C_OK_256; C_WARN=$C_WARN_256; C_CRIT=$C_CRIT_256;;
  basic) c_empty=$EMPTY_16; C_OK=$C_OK_16; C_WARN=$C_WARN_16; C_CRIT=$C_CRIT_16;;
  none) C_OK="0"; C_WARN="0"; C_CRIT="0";;
esac
# bar: $1 = so o day, $2 = ma mau theo nguong. In "[" + o day + o rong + "]".
bar() {
  local filled=$1 col=$2 i full="" emp=""
  (( filled > 10 )) && filled=10
  if [[ $cmode == none ]]; then
    for i in {1..10}; do (( i <= filled )) && full+="$BAR_FULL" || full+="$BAR_EMPTY"; done
    print -rn -- "[$full]"; return
  fi
  for i in {1..10}; do
    (( i > filled )) && break
    if [[ $cmode == basic ]]; then full+="$BAR_FULL"; else full+="${esc}[${grad[i]}m${BAR_FULL}"; fi
  done
  for i in {1..10}; do (( i > filled )) && emp+="$BAR_EMPTY"; done
  [[ $cmode == basic && -n "$full" ]] && full="${esc}[${col}m${full}"
  [[ -n "$full" ]] && full+="${esc}[0m"
  [[ -n "$emp" ]] && emp="${esc}[${c_empty}m${emp}${esc}[0m"
  print -rn -- "$(paint $C_SEP '[')${full}${emp}$(paint $C_SEP ']')"
}
lvlcol() { case $1 in 2) print -rn -- $C_CRIT;; 1) print -rn -- $C_WARN;; *) print -rn -- $C_OK;; esac }
# fit: $1 = chuoi co ma ANSI, $2 = do rong toi da.
# Vuot thi cat o diem hien thi, bo ca cum thanh [..] neu cat giua thanh, them reset ANSI + "...".
fit() {
  local s="$1" max=$2 plain out="" ch vis=0 keep i n=${#1} inbar=0 barstart=0
  plain=${s//${esc}\[[0-9;]#m/}
  if (( ${#plain} <= max )); then print -r -- "$s"; return; fi
  if (( max <= 3 )); then print -r -- "${(l:$max::.:)}"; return; fi
  keep=$(( max - 3 ))
  for (( i=1; i<=n; i++ )); do
    ch=${s[i]}
    if [[ "$ch" == "$esc" ]]; then
      while (( i <= n )) && [[ "${s[i]}" != m ]]; do out+=${s[i]}; (( i++ )); done
      out+=m; continue
    fi
    (( vis >= keep )) && break
    if [[ "$ch" == "[" ]]; then inbar=1; barstart=${#out}; fi
    [[ "$ch" == "%" || "$ch" == "-" ]] && inbar=0
    out+=$ch; (( vis++ ))
  done
  (( inbar )) && out=${out[1,$barstart]}
  out=${out%%( )##}
  if [[ $cmode == none ]]; then print -r -- "${out}..."; else print -r -- "${out}${esc}[0m..."; fi
}
vlen() { local p=${1//$'\e'\[[0-9;]#m/}; print -r -- ${#p}; }
# f: in "lam_tron|so_o_day|muc|resets_at" (rong neu thieu). Tinh bang jq.
f() {
  print -r -- "$input" | jq -r --arg w "$1" '
    .rate_limits[$w] // empty | select((.used_percentage | type) == "number")
    | (if .used_percentage < 0 then 0 elif .used_percentage > 100 then 100 else .used_percentage end) as $p
    | "\($p | round)|\(($p / 10) | floor)|\(if $p >= 90 then 2 elif $p >= 70 then 1 else 0 end)|\(.resets_at // "")"' 2>/dev/null
}
# seg: $1 = nhan, $2 = ket qua f, $3 = format ngay gio, $4 = short | nobar
seg() {
  local v="$2" pct filled lvl ts rest col out nl
  nl=$(paint $C_NEUTRAL "$1")
  if [[ -z "$v" ]]; then
    if [[ "$4" == nobar ]]; then print -rn -- "$nl --"; else print -rn -- "$nl $(bar 0 $C_OK) --"; fi
    return
  fi
  pct=${v%%|*}; rest=${v#*|}
  filled=${rest%%|*}; rest=${rest#*|}
  lvl=${rest%%|*}; ts=${rest#*|}
  col=$(lvlcol $lvl)
  if [[ "$4" == nobar ]]; then
    print -rn -- "$nl $(paint $col "${pct}%")"; return
  fi
  out="$nl $(bar $filled $col) $(paint $col "${pct}%")"
  if [[ "$ts" == <-> && -z "$4" ]]; then
    out="$out $(paint $C_SEP "(reset $(date -r "$ts" +"$3" 2>/dev/null))")"
  fi
  print -rn -- "$out"
}
jqs() { print -r -- "$input" | jq -r "$1" 2>/dev/null; }
model=$(jqs '.model.display_name | select(type == "string" and length > 0)')
effort=$(jqs '.effort.level | select(type == "string" and length > 0)')
# c: "lam_tron|so_o|muc|chi_tiet" ; chi_tiet = "64k/200k" neu co du so lieu
c=$(print -r -- "$input" | jq -r '
  .context_window // empty | select((.used_percentage | type) == "number")
  | (if .used_percentage < 0 then 0 elif .used_percentage > 100 then 100 else .used_percentage end) as $p
  | (if (.total_input_tokens | type) == "number" and (.context_window_size | type) == "number" and .context_window_size > 0
     then "\((.total_input_tokens / 1000) | round)k/\((.context_window_size / 1000) | round)k" else "" end) as $d
  | (if (.total_output_tokens | type) == "number" then " | out \((.total_output_tokens / 1000) | round)k" else "" end) as $o
  | (if (.current_usage.cache_read_input_tokens | type) == "number" and (.current_usage.cache_creation_input_tokens | type) == "number"
     then " | cache \(((.current_usage.cache_read_input_tokens + .current_usage.cache_creation_input_tokens) / 1000) | round)k" else "" end) as $k
  | (if $d == "" then "" else $d + $o + $k end) as $d
  | "\($p | round)|\(($p / 10) | floor)|\(if $p >= 90 then 2 elif $p >= 70 then 1 else 0 end)|\($d)"' 2>/dev/null)
# ctxseg: $1 = ket qua c, $2 = pad | full | short | nodet | nobar | (pad = full + dem nhan).
ctxseg() {
  local v="$1" pct filled lvl det rest out col lp=""
  [[ "$2" == pad ]] && lp=$lpad
  if [[ -z "$v" ]]; then
    if [[ "$2" == nobar ]]; then print -rn -- "$(paint $C_NEUTRAL Ctx) --"; else print -rn -- "$(paint $C_NEUTRAL Context:)$lp $(bar 0 $C_OK) --"; fi
    return
  fi
  pct=${v%%|*}; rest=${v#*|}
  filled=${rest%%|*}; rest=${rest#*|}
  lvl=${rest%%|*}; det=${rest#*|}
  col=$(lvlcol $lvl)
  if [[ "$2" == nobar ]]; then
    print -rn -- "$(paint $C_NEUTRAL Ctx) $(paint $col "${pct}%")"; return
  fi
  out="$(paint $C_NEUTRAL Context:)$lp $(bar $filled $col) $(paint $col "${pct}%")"
  [[ "$2" == short ]] && det=${det%% | *}
  [[ "$2" == nodet ]] && det=""
  if [[ -n "$det" ]]; then
    if [[ $cmode == none ]]; then out="$out ($det)"; else out="$out $(paint $C_SEP "($det)")"; fi
  fi
  print -rn -- "$out"
}
cols=${COLUMNS:-80}
[[ "$cols" == <-> && $cols -gt 0 ]] || cols=80

# ===== Dong 1: Context | Model | Effort =====
mdlv=$(paint $C_MODEL "${model:---}")
case "${effort:l}" in
  low) ec=$C_EFFORT_LOW;; medium) ec=$C_EFFORT_MEDIUM;; high) ec=$C_EFFORT_HIGH;;
  xhigh) ec=$C_EFFORT_XHIGH;; max) ec=$C_EFFORT_MAX;; *) ec=$C_EFFORT_OTHER;;
esac
# l2: $1 = muc 1..8
l2() {
  local mdl="$(paint $C_NEUTRAL Model:) $mdlv" eff="$(paint $C_NEUTRAL Effort:) $(paint $ec "${effort:---}")"
  case $1 in
    1) print -rn -- "$(ctxseg "$c" pad)$sep$mdl$sep$eff";;
    2) print -rn -- "$(ctxseg "$c" full)$sep$mdl$sep$eff";;
    3) print -rn -- "$(ctxseg "$c" short)$sep$mdl$sep$eff";;
    4) print -rn -- "$(ctxseg "$c" nodet)$sep$mdl$sep$eff";;
    5) print -rn -- "$(ctxseg "$c" nodet)$sep$mdl";;
    6) print -rn -- "$(ctxseg "$c" nodet)$sep$mdlv";;
    7) print -rn -- "$(ctxseg "$c" nobar)$sep$mdlv";;
    *) ctxseg "$c" nobar;;
  esac
}
# ===== Dong 2: Quota =====
f5=$(f five_hour); f7=$(f seven_day)
# l3: $1 = muc 1..4
l3f() {
  local q="$(paint $C_NEUTRAL Quota:)"
  case $1 in
    1) print -rn -- "$q $(seg 5-Hours "$f5" '%H:%M')$sep$(seg Week "$f7" '%d/%m %H:%M')";;
    2) print -rn -- "$q $(seg 5-Hours "$f5" '%H:%M' short)$sep$(seg Week "$f7" '%d/%m %H:%M' short)";;
    3) print -rn -- "$q $(seg 5h "$f5" '%H:%M' short)$sep$(seg Wk "$f7" '%d/%m %H:%M' short)";;
    *) print -rn -- "$(paint $C_NEUTRAL Quota) $(seg 5h "$f5" '' nobar)$sep$(seg Wk "$f7" '' nobar)";;
  esac
}
# Rut gon dong 1: bo dem, bo out|cache, bo (used/size), bo Effort, bo nhan "Model:", bo thanh (Ctx NN% | model), bo model, cat "...".
l1=""; found=0
for lv in 1 2 3 4 5 6 7 8; do
  l1=$(l2 $lv)
  if (( $(vlen "$l1") <= cols )); then found=1; break; fi
done
if (( ! found )); then l1=$(l2 8); fi
fit "$l1" $cols

l3=""; found=0
for lv in 1 2 3 4; do
  l3=$(l3f $lv)
  if (( $(vlen "$l3") <= cols )); then found=1; break; fi
done
if (( ! found )); then l3=$(l3f 4); fi
fit "$l3" $cols
