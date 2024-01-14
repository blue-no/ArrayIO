import gc
import shutil

import memory_profiler

from arrayio import Array, read_csv_array, set_verbosity
from tests import DIR, LINES, VALUES

fp_in = DIR / "input.csv"
fp_base = DIR / "output_base.csv"
fp_out1 = DIR / "output1.csv"
fp_out2 = DIR / "output2.csv"

set_verbosity(True)


@memory_profiler.profile
def test_read_csv():
    print("\n" + test_read_csv.__name__)
    array = read_csv_array(
        fp=fp_in,
        range_=LINES,
        lazy=False,
    )
    assert array.get_values() == VALUES
    gc.collect()


@memory_profiler.profile
def test_write_csv():
    print("\n" + test_write_csv.__name__)
    shutil.copyfile(fp_base, fp_out1)
    Array(values=VALUES).write_to_csv(fp=fp_out1)
    gc.collect()


@memory_profiler.profile
def test_csv_io_lazy():
    print("\n" + test_csv_io_lazy.__name__)
    shutil.copyfile(fp_base, fp_out2)

    array = read_csv_array(
        fp=fp_in,
        range_=LINES,
        lazy=True,
    )

    array.write_to_csv(fp=fp_out2)

    assert array.get_values() == VALUES

    array.write_to_csv(fp=fp_out2)
    gc.collect()
