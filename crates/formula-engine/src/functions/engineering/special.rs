use crate::value::ErrorKind;

pub(crate) fn erf(x: f64) -> Result<f64, ErrorKind> {
    if !x.is_finite() {
        return Err(ErrorKind::Num);
    }
    let out = libm::erf(x);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub(crate) fn erfc(x: f64) -> Result<f64, ErrorKind> {
    if !x.is_finite() {
        return Err(ErrorKind::Num);
    }
    let out = libm::erfc(x);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub(crate) fn besselj(x: f64, n: i32) -> Result<f64, ErrorKind> {
    if !x.is_finite() {
        return Err(ErrorKind::Num);
    }
    if n < 0 {
        return Err(ErrorKind::Num);
    }
    let out = libm::jn(n, x);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub(crate) fn bessely(x: f64, n: i32) -> Result<f64, ErrorKind> {
    if !x.is_finite() {
        return Err(ErrorKind::Num);
    }
    if n < 0 {
        return Err(ErrorKind::Num);
    }
    // Excel's BESSELY is restricted to x > 0; negative values yield #NUM! (complex result).
    if x <= 0.0 {
        return Err(ErrorKind::Num);
    }
    let out = libm::yn(n, x);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub(crate) fn besseli(x: f64, n: i32) -> Result<f64, ErrorKind> {
    if !x.is_finite() {
        return Err(ErrorKind::Num);
    }
    if n < 0 {
        return Err(ErrorKind::Num);
    }

    // Excel defines BESSELI for real x and integer n. For negative x, I_n(-x)=(-1)^n I_n(x).
    let ax = x.abs();
    let mut out = match n {
        0 => besseli0(ax),
        1 => besseli1(ax),
        _ => {
            if ax == 0.0 {
                return Ok(0.0);
            }
            let mut i_nm1 = besseli0(ax);
            let mut i_n = besseli1(ax);
            for k in 1..n {
                let kf = k as f64;
                let i_np1 = i_nm1 - (2.0 * kf / ax) * i_n;
                i_nm1 = i_n;
                i_n = i_np1;
            }
            i_n
        }
    };
    if x < 0.0 && (n % 2 != 0) {
        out = -out;
    }

    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

pub(crate) fn besselk(x: f64, n: i32) -> Result<f64, ErrorKind> {
    if !x.is_finite() {
        return Err(ErrorKind::Num);
    }
    if n < 0 {
        return Err(ErrorKind::Num);
    }
    // Excel's BESSELK is restricted to x > 0.
    if x <= 0.0 {
        return Err(ErrorKind::Num);
    }

    let out = match n {
        0 => besselk0(x),
        1 => besselk1(x),
        _ => {
            let mut k_nm1 = besselk0(x);
            let mut k_n = besselk1(x);
            for k in 1..n {
                let kf = k as f64;
                let k_np1 = k_nm1 + (2.0 * kf / x) * k_n;
                k_nm1 = k_n;
                k_n = k_np1;
            }
            k_n
        }
    };

    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

// -----------------------------------------------------------------------------
// Modified Bessel I0/I1/K0/K1
// -----------------------------------------------------------------------------
// Excel's BESSELI/BESSELK correspond to the modified Bessel functions I_n and K_n.
//
// `libm` does not currently expose I/K helpers, so we use standard polynomial
// approximations (from classic Cephes/Numerical Recipes implementations) for
// orders 0 and 1, then apply the integer-order recurrence.

fn besseli0(x: f64) -> f64 {
    let ax = x.abs();
    if ax < 3.75 {
        let t = ax / 3.75;
        let t2 = t * t;
        1.0 + t2
            * (3.515_622_9
                + t2 * (3.089_942_4
                    + t2 * (1.206_749_2
                        + t2 * (0.265_973_2 + t2 * (0.036_076_8 + t2 * 0.004_581_3)))))
    } else {
        let t = 3.75 / ax;
        let poly = 0.398_942_28
            + t * (0.013_285_92
                + t * (0.002_253_19
                    + t * (-0.001_575_65
                        + t * (0.009_162_81
                            + t * (-0.020_577_06
                                + t * (0.026_355_37 + t * (-0.016_476_33 + t * 0.003_923_77)))))));
        (ax.exp() / ax.sqrt()) * poly
    }
}

fn besseli1(x: f64) -> f64 {
    let ax = x.abs();
    let out = if ax < 3.75 {
        let t = ax / 3.75;
        let t2 = t * t;
        ax * (0.5
            + t2 * (0.878_905_94
                + t2 * (0.514_988_69
                    + t2 * (0.150_849_34
                        + t2 * (0.026_587_33 + t2 * (0.003_015_32 + t2 * 0.000_324_11))))))
    } else {
        let t = 3.75 / ax;
        let poly = 0.398_942_28
            + t * (-0.039_880_24
                + t * (-0.003_620_18
                    + t * (0.001_638_01
                        + t * (-0.010_315_55
                            + t * (0.022_829_67
                                + t * (-0.028_953_12 + t * (0.017_876_54 - t * 0.004_200_59)))))));
        (ax.exp() / ax.sqrt()) * poly
    };

    if x < 0.0 {
        -out
    } else {
        out
    }
}

fn besselk0(x: f64) -> f64 {
    if x <= 0.0 {
        return f64::NAN;
    }
    if x <= 2.0 {
        let t = (x * x) / 4.0;
        let poly = -0.577_215_66
            + t * (0.422_784_20
                + t * (0.230_697_56
                    + t * (0.034_885_90
                        + t * (0.002_626_98 + t * (0.000_107_50 + t * 0.000_007_40)))));
        -((x / 2.0).ln()) * besseli0(x) + poly
    } else {
        let t = 2.0 / x;
        let poly = 1.253_314_14
            + t * (-0.078_323_58
                + t * (0.021_895_68
                    + t * (-0.010_624_46
                        + t * (0.005_878_72 + t * (-0.002_515_40 + t * 0.000_532_08)))));
        (-(x)).exp() / x.sqrt() * poly
    }
}

fn besselk1(x: f64) -> f64 {
    if x <= 0.0 {
        return f64::NAN;
    }
    if x <= 2.0 {
        let t = (x * x) / 4.0;
        let poly = 1.0
            + t * (0.154_431_44
                + t * (-0.672_785_79
                    + t * (-0.181_568_97
                        + t * (-0.019_194_02 + t * (-0.001_104_04 + t * (-0.000_046_86))))));
        ((x / 2.0).ln()) * besseli1(x) + poly / x
    } else {
        let t = 2.0 / x;
        let poly = 1.253_314_14
            + t * (0.234_986_19
                + t * (-0.036_556_20
                    + t * (0.015_042_68
                        + t * (-0.007_803_53 + t * (0.003_256_14 + t * (-0.000_682_45))))));
        (-(x)).exp() / x.sqrt() * poly
    }
}
