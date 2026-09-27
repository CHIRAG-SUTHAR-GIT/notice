from src.csv_fixer import fix_csv_text


def test_fix_csv_text_repairs_extra_commas_in_second_column():
    row = "1,ACME BANK, MAIN BRANCH,1,2,3,4,5,6,7,8,9,10,11,12,13,14"

    fixed, stats = fix_csv_text(row)

    assert fixed == '1,"ACME BANK, MAIN BRANCH",1,2,3,4,5,6,7,8,9,10,11,12,13,14'
    assert stats.fixed_lines == 1
    assert stats.unchanged_lines == 0


def test_fix_csv_text_keeps_aligned_rows_unchanged():
    row = "1,ACME BANK MAIN BRANCH,1,2,3,4,5,6,7,8,9,10,11,12,13,14"

    fixed, stats = fix_csv_text(row)

    assert fixed == row
    assert stats.fixed_lines == 0
    assert stats.unchanged_lines == 1
