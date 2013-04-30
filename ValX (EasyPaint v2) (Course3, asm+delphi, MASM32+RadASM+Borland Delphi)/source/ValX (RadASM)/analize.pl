open(IN, 'Добавить.txt') || die "cannot read: $!";
$c=0;
$a=0;
$q=0;
while (<IN>)
{
	if (/^\+/)
	{
		$c++;
		$a++;
	}
	elsif (/^\-/)
	{
		$c++;
	}
	elsif (/^\?/)
	{
		$q++;
	}
}

close(IN);

if ($c) {$pers = $a/$c*100;}
else {$pers = '<не вичислимо>';}

open(OUT, '>Res.txt') || die "cannot write: $!";
print OUT <<END
Всего заданий      : $c
Выполнено          : $a
Процент выполнения : $pers%
Вопросов           : $q
END
;
close(OUT);