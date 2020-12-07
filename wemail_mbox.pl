# ========== WeMail → mbox 変換ツール ==========

  use strict;
  use warnings;
  use IO::Handle;

  our ($kgr);

  if (@ARGV <= 0) {
    $kgr = '\\';
  }
  elsif (@ARGV == 1) {
    $kgr = $ARGV[0];
  }
  else {
    print STDERR 'Usage: ' . $0 . " 区切り文字\n";
    exit 1;
  }

  &cur('.');


# $_[0] 配下(サブフォルダも)の *.wem を変換して mbox にする
sub cur {
  my ($cwd) = @_;
  my ($file, $wem, $dh);

  $wem = 0;
  $dh = IO::Handle -> new();
  unless (opendir $dh, $cwd) {
    warn "フォルダ $cwd を開けません\n";
    return;
  }
  while ($file = readdir $dh) {
    next if ($file eq '.' || $file eq '..');
    if (-d $cwd . '\\' . $file) {
      &cur($cwd . '\\' . $file);
    }
    elsif (substr($file, -4) eq '.wem') {
      $wem = 1;
    }
  }
  closedir $dh;
  if ($wem) {
    my ($res, $mf, $fpr);
    ($res, $mf) = (&mbox($cwd), $cwd . '\\mbox');
    if ($mf =~ /^\.?\\(.*)/) { $mf = $1; }

    { # フォルダの区切り文字を変換
      my ($str);
      $str = $mf;
      $mf = '';
      while ($str ne '') {
        my ($c, $code);
        $c = substr($str, 0, 1);
        ($str, $code) = (substr($str, 1), ord($c));
        if (0x80 <= $code && $code < 0xa0 || 0xe0 <= $code) {
          if ($str ne '') {
            $c .= substr($str, 0, 1);
            $str = substr($str, 1);
          }
        }
        $c = $kgr if ($c eq '\\');
        $mf .= $c;
      }
    }

    $fpr = IO::Handle -> new();
    unless (open $fpr, '> ' . $mf) {
      warn "ファイル $mf の書き込みに失敗しました\n";
      return;
    }
    print $fpr $res;
    close $fpr;
  }

}


# $_[0] 直下の *.wem を変換して mbox にする
sub mbox {
  my ($cwd) = @_;
  my (@lt, $fil, $otz, $dt, $dhw);
  local our ($cy, %zen, $oy, $omo, $od, $oh, $omi, $os, $oyb, $ofr);

  @lt = localtime;
  $cy = $lt[5] + 1900;
  %zen = ();

  ($otz, $oy, $omo, $od, $oh) = (0, $cy, $lt[4] + 1, $lt[3], $lt[2]);
  ($omi, $os, $oyb, $ofr) = ($lt[1], $lt[0], $lt[6], 'daemon@localhost');

  $dhw = IO::Handle -> new();
  unless (opendir $dhw, $cwd) {
    warn "フォルダ $cwd を開けません\n";
    return '';
  }
  while ($fil = readdir $dhw) {
    next if ($fil eq '.' || $fil eq '..' || -d $cwd . '\\' . $fil);
    if (substr($fil, -4) eq '.wem') {
      my ($naiyou, $buf, $yb, $from, $meid, $myb, $fpi);

      $fil = $cwd . '\\' . $fil;

      $fpi = IO::Handle -> new();
      unless (open $fpi, '< ' . $fil) {
        warn "ファイル $fil の読み出しに失敗しました\n";
        next;
      }

      $from = $dt = $naiyou = $yb = $meid = $myb = '';
      while ($buf = <$fpi>) {
        chomp $buf;

        if ($buf =~ /^Date: / && $dt eq '') {
          my ($h, $tz, $d, $y, $mo, $mi, $s);
          if ($buf =~
            / (\d+) (\w+) (\d\d\d\d) (\d+):(\d+):(\d+) +([\+\-]\d\d)(\d\d)/) {
            my ($i, $tzmin);
            $i = index("JanFebMarAprMayJunJulAugSepOctNovDec", $2);
            $tzmin = $7 * 60;
            if (substr($7, 0, 1) eq '-') {
              $tzmin -= $8;
            }
            elsif (substr($7, 0, 1) eq '+') {
              $tzmin += $8;
            }
            ($tz, $y) = (9 * 60 - $tzmin, $3);
            while ($y < $cy - 50) { $y += 100; }
            $mo = $i >= 0 ? $i / 3 + 1 : $omo;
            ($d, $h, $mi, $s) = ($1, $4, $5, $6);
            $yb = &youbi($y, $mo, $d);
            if ($buf =~
              /( \d+ \w+ )(\d\d\d\d)( \d+:\d+:\d+ +[\+\-]\d\d\d\d)/) {
              $buf = $` . $1 . $y . $3 . $';
            }
          }
          elsif ($buf =~
            / (\d+) (\w+) (\d+) (\d+):(\d+):(\d+) GMT$/) {
            my ($i, $fix);
            $i = index("JanFebMarAprMayJunJulAugSepOctNovDec", $2);
            ($tz, $fix, $y) = (9 * 60, 0, $3);
            if (length($y) != 4) {
              while ($y < $cy - 50) { $y += 100; }
              while ($y >= $cy + 50) { $y -= 100; }
            }
            while ($y < $cy - 50) {
              $y += 100;
              $fix = 1;
            }
            $mo = $i >= 0 ? $i / 3 + 1 : $omo;
            ($d, $h, $mi, $s) = ($1, $4, $5, $6);
            $yb = &youbi($y, $mo, $d);
            if ($fix) {
              if ($buf =~
                /( \d+ \w+ )(\d+)( \d+:\d+:\d+ GMT)$/) {
                $buf = $` . $1 . $y . $3;
              }
            }
          }
          elsif ($buf =~
            / (\d+) (\w+) (\d\d\d\d) (\d+):(\d+) +([\+\-]\d\d)(\d\d)$/) {
            my ($i, $tzmin);
            $i = index("JanFebMarAprMayJunJulAugSepOctNovDec", $2);
            $tzmin = $6 * 60;
            if (substr($6, 0, 1) eq '-') {
              $tzmin -= $7;
            }
            elsif (substr($6, 0, 1) eq '+') {
              $tzmin += $7;
            }
            ($tz, $y) = (9 * 60 - $tzmin, $3);
            while ($y < $cy - 50) { $y += 100; }
            $mo = $i >= 0 ? $i / 3 + 1 : $omo;
            ($d, $h, $mi, $s) = ($1, $4, $5, 0);
            $yb = &youbi($y, $mo, $d);
            if ($buf =~
              /( \d+ \w+ )(\d\d\d\d)( \d+:\d+ +[\+\-]\d\d\d\d)$/) {
              $buf = $` . $1 . $y . $3;
            }
          }
          elsif ($buf =~
            / (\d+) (\w+) (\d+) (\d+):(\d+):(\d+) ?$/) {
            my ($i, $fix);
            $i = index("JanFebMarAprMayJunJulAugSepOctNovDec", $2);
            ($tz, $fix, $y) = (0, 0, $3);
            if (length($y) != 4) {
              while ($y < $cy - 50) { $y += 100; }
              while ($y >= $cy + 50) { $y -= 100; }
            }
            while ($y < $cy - 50) {
              $y += 100;
              $fix = 1;
            }
            $mo = $i >= 0 ? $i / 3 + 1 : $omo;
            ($d, $h, $mi, $s) = ($1, $4, $5, $6);
            $yb = &youbi($y, $mo, $d);
            if ($buf =~ /( \d+ \w+ )(\d+)( \d+:\d+:\d+ )$/) {
              $buf = $` . $1 . ($fix ? $y : $2) . $3 . '+0900';
            }
          }
          elsif ($buf =~
            / , (\d+) (\w\w\w)(\d\d\d\d) (\d\d):(\d\d):(\d\d) 1100$/) {
            my ($i);
            $i = index("JAN FEB MAR APR MAY JUN JUL AUG SEP OCT NOV DEC", $2);
            ($tz, $y) = (60, $3);
            $mo = $i >= 0 ? $i / 4 + 1 : $omo;
            ($d, $h, $mi, $s) = ($1, $4, $5, $6);
            $yb = &youbi($y, $mo, $d);
            if ($buf =~
              / , \d+ \w\w\w\d\d\d\d \d\d:\d\d:\d\d 1100$/) {
              $buf = $` . ' ';
              $buf .= substr('SunMonTueWedThuFriSat', $yb * 3, 3) . ', ';
              $buf .= ($d + 0) . ' ';
              $buf .= substr("JanFebMarAprMayJunJulAugSepOctNovDec",
                ($mo - 1) * 3, 3) . ' ';
              $buf .= sprintf("%04d %02d:%02d:%02d +0900",
                $y, $h, $mi, $s);
            }
          }
          else {
            ($tz, $y, $mo, $d) = ($otz, $oy, $omo, $od);
            ($h, $mi, $s, $yb) = ($oh, $omi, $os, $oyb);
          }

          ($otz, $oy, $omo, $od) = ($tz, $y, $mo, $d);
          ($oh, $omi, $os, $oyb) = ($h, $mi, $s, $yb);

          $mi += $tz;
          while ($mi >= 60) {
            $h ++;
            $mi -= 60;
          }
          while ($mi < 0) {
            $h --;
            $mi += 60;
          }
          if ($h < 0) {
            $h += 24;
            ($y, $mo, $d, $yb) = &yesterday($y, $mo, $d, $yb);
          }
          elsif ($h > 23) {
            $h -= 24;
            ($y, $mo, $d, $yb) = &tomorrow($y, $mo, $d, $yb);
          }
          $dt = sprintf("%04d%02d%02d%02d%02d%02d",
            $y, $mo, $d, $h, $mi, $s);
        }
        elsif ($buf =~ /^From: ?/i && ($from eq '')) {
          my ($fr);
          $fr = $';

          if ($buf =~ /^From: ".*" <(.+)>$/i) {
            $from = $1;
          }
          elsif ($buf =~ /^From:.*<(.+)>$/i) {
            $from = $1;
          }
          else {
            $from = $fr;
          }
          $ofr = $from;
        }
        elsif ($buf =~ /^From /) {
          if ($naiyou eq '') {
            $from = $buf;
            $ofr = $from;
            next;
          }
          $buf = '>' . $buf;
        }
        elsif ($buf =~ /^Message-Id: </) {
          my ($mei);
          $mei = $';
          if ($mei =~ /^(\d\d\d\d)(\d\d)(\d\d)(\d\d)(\d\d)(\d\d)/) {
            my ($h, $d, $y, $mo, $mi, $s);
            ($y, $mo, $d, $h, $mi, $s) = ($1, $2, $3, $4, $5, $6);
            $meid = $y . $mo . $d . $h . $mi . $s;
            $myb = &youbi($y, $mo, $d);
          }
        }
        $naiyou .= $buf . "\n";
      }

      close $fpi;

      $dt = $meid if ($dt eq '');
      $yb = $myb if ($yb eq '');
      &bun($naiyou, $dt, $from, $yb);

    }
  }
  closedir $dhw;

  # 結果文字列を生成
  $fil = '';
  foreach $dt (sort keys %zen) { $fil .= ($zen{$dt} . "\n"); }

  return $fil;

  sub bun {
    my ($naiyou, $dt, $from, $yb) = @_;
    my ($cdt, @lt, $fhed, $fh);

    if ($dt eq '') {
      $dt = sprintf("%04d%02d%02d%02d%02d%02d",
        $oy, $omo, $od, $oh, $omi, $os);
    }

    unless ($from) { $from = $ofr; }

    if ($yb eq '') { $yb = $oyb; }

    $fh = 'From ' . $from . ' ';
    $fh .= substr("SunMonTueWedThuFriSat", $yb * 3, 3) . ' ';
    $fh .= substr("JanFebMarAprMayJunJulAugSepOctNovDec",
      (substr($dt, 4, 2) - 1) * 3, 3) . ' ';
    $fh .= sprintf("%2d %02d:%02d:%02d %04d", substr($dt, 6, 2),
      substr($dt, 8, 2), substr($dt, 10, 2),
      substr($dt, 12, 2), substr($dt, 0, 4));

    while ($zen{$dt}) { $dt .= 'a'; }
    $zen{$dt} = $fh . "\n" . $naiyou;

  }

  sub youbi {
    my ($y, $m, $d) = @_;
    my ($c, $w);

    if ($m <= 2) {
      $m += 12;
      $y --;
    }
    $c = int($y / 100);
    $y = $y % 100;

    $w = int($c / 4) - 2 * $c + int($y / 4)
      + $y + int(26 * ($m + 1) / 10) + $d;

    return ($w + 6) % 7;    # 日:0 土:6
  }

  sub tomorrow {
    my ($y, $m, $d, $w) = @_;
    $y = $oy if ($y eq '');
    $m = $omo if ($m eq '');
    $d = $od if ($d eq '');
    $w = $oyb if ($w eq '');
    $w = ($w + 1) % 7;
    $d ++;
    if ($d == 29) {
      if ($m == 2) {
        if ($y % 4) {
          $m ++;
          $d = 1;
        }
        elsif ($y % 100) {
        }
        elsif ($y % 400) {
          $m ++;
          $d = 1;
        }
      }
    }
    elsif ($d == 30) {
      if ($m == 2) {
        $m ++;
        $d = 1;
      }
    }
    elsif ($d == 31) {
      if ($m == 2 || $m == 4 || $m == 6 || $m == 9 || $m == 11) {
        $m ++;
        $d = 1;
      }
    }
    elsif ($d == 32) {
      $m ++;
      $d = 1;
    }
    if ($m > 12) {
      $y ++;
      while ($y >= $cy + 50) { $y -= 100; }
      $m = 1;
    }
    return ($y, $m, $d, $w);
  }

  sub yesterday {
    my ($y, $m, $d, $w) = @_;
    $y = $oy if ($y eq '');
    $m = $omo if ($m eq '');
    $d = $od if ($d eq '');
    $w = $oyb if ($w eq '');
    $w = ($w + 6) % 7;
    $d --;
    if ($d <= 0) {
      if ($m == 1 || $m == 2 || $m == 4 || $m == 6 || $m == 8 ||
          $m == 9 || $m == 11) {
        $m --;
        $d = 31;
      }
      elsif ($m == 3) {
        $m --;
        if ($y % 4) {
          $d = 28;
        }
        elsif ($y % 100) {
          $d = 29;
        }
        elsif ($y % 400) {
          $d = 28;
        }
        else {
          $d = 29;
        }
      }
      else {
        $m --;
        $d = 30;
      }
      if ($m <= 0) {
        $y --;
        while ($y < $cy - 50) { $y += 100; }
        $m = 12;
      }
    }
    return ($y, $m, $d, $w);
  }

}

__END__
