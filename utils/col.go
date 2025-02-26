package utils

import "github.com/xuri/excelize/v2"

// 从 0 开始
var ColumnNameToNumberCache = map[string]int{
	// 大写字母 A-Z
	"A": 0, "B": 1, "C": 2, "D": 3, "E": 4, "F": 5, "G": 6, "H": 7, "I": 8, "J": 9,
	"K": 10, "L": 11, "M": 12, "N": 13, "O": 14, "P": 15, "Q": 16, "R": 17, "S": 18,
	"T": 19, "U": 20, "V": 21, "W": 22, "X": 23, "Y": 24, "Z": 25,
	// 大写字母 AA-EZ
	"AA": 26, "AB": 27, "AC": 28, "AD": 29, "AE": 30, "AF": 31, "AG": 32, "AH": 33,
	"AI": 34, "AJ": 35, "AK": 36, "AL": 37, "AM": 38, "AN": 39, "AO": 40, "AP": 41,
	"AQ": 42, "AR": 43, "AS": 44, "AT": 45, "AU": 46, "AV": 47, "AW": 48, "AX": 49,
	"AY": 50, "AZ": 51,
	"BA": 52, "BB": 53, "BC": 54, "BD": 55, "BE": 56, "BF": 57, "BG": 58, "BH": 59,
	"BI": 60, "BJ": 61, "BK": 62, "BL": 63, "BM": 64, "BN": 65, "BO": 66, "BP": 67,
	"BQ": 68, "BR": 69, "BS": 70, "BT": 71, "BU": 72, "BV": 73, "BW": 74, "BX": 75,
	"BY": 76, "BZ": 77,
	"CA": 78, "CB": 79, "CC": 80, "CD": 81, "CE": 82, "CF": 83, "CG": 84, "CH": 85,
	"CI": 86, "CJ": 87, "CK": 88, "CL": 89, "CM": 90, "CN": 91, "CO": 92, "CP": 93,
	"CQ": 94, "CR": 95, "CS": 96, "CT": 97, "CU": 98, "CV": 99, "CW": 100, "CX": 101,
	"CY": 102, "CZ": 103,
	"DA": 104, "DB": 105, "DC": 106, "DD": 107, "DE": 108, "DF": 109, "DG": 110, "DH": 111,
	"DI": 112, "DJ": 113, "DK": 114, "DL": 115, "DM": 116, "DN": 117, "DO": 118, "DP": 119,
	"DQ": 120, "DR": 121, "DS": 122, "DT": 123, "DU": 124, "DV": 125, "DW": 126, "DX": 127,
	"DY": 128, "DZ": 129,
	"EA": 130, "EB": 131, "EC": 132, "ED": 133, "EE": 134, "EF": 135, "EG": 136, "EH": 137,
	"EI": 138, "EJ": 139, "EK": 140, "EL": 141, "EM": 142, "EN": 143, "EO": 144, "EP": 145,
	"EQ": 146, "ER": 147, "ES": 148, "ET": 149, "EU": 150, "EV": 151, "EW": 152, "EX": 153,
	"EY": 154, "EZ": 155,

	// 小写字母 a-z
	"a": 0, "b": 1, "c": 2, "d": 3, "e": 4, "f": 5, "g": 6, "h": 7, "i": 8, "j": 9,
	"k": 10, "l": 11, "m": 12, "n": 13, "o": 14, "p": 15, "q": 16, "r": 17, "s": 18,
	"t": 19, "u": 20, "v": 21, "w": 22, "x": 23, "y": 24, "z": 25,
	// 小写字母 aa-ez
	"aa": 26, "ab": 27, "ac": 28, "ad": 29, "ae": 30, "af": 31, "ag": 32, "ah": 33,
	"ai": 34, "aj": 35, "ak": 36, "al": 37, "am": 38, "an": 39, "ao": 40, "ap": 41,
	"aq": 42, "ar": 43, "as": 44, "at": 45, "au": 46, "av": 47, "aw": 48, "ax": 49,
	"ay": 50, "az": 51,
	"ba": 52, "bb": 53, "bc": 54, "bd": 55, "be": 56, "bf": 57, "bg": 58, "bh": 59,
	"bi": 60, "bj": 61, "bk": 62, "bl": 63, "bm": 64, "bn": 65, "bo": 66, "bp": 67,
	"bq": 68, "br": 69, "bs": 70, "bt": 71, "bu": 72, "bv": 73, "bw": 74, "bx": 75,
	"by": 76, "bz": 77,
	"ca": 78, "cb": 79, "cc": 80, "cd": 81, "ce": 82, "cf": 83, "cg": 84, "ch": 85,
	"ci": 86, "cj": 87, "ck": 88, "cl": 89, "cm": 90, "cn": 91, "co": 92, "cp": 93,
	"cq": 94, "cr": 95, "cs": 96, "ct": 97, "cu": 98, "cv": 99, "cw": 100, "cx": 101,
	"cy": 102, "cz": 103,
	"da": 104, "db": 105, "dc": 106, "dd": 107, "de": 108, "df": 109, "dg": 110, "dh": 111,
	"di": 112, "dj": 113, "dk": 114, "dl": 115, "dm": 116, "dn": 117, "do": 118, "dp": 119,
	"dq": 120, "dr": 121, "ds": 122, "dt": 123, "du": 124, "dv": 125, "dw": 126, "dx": 127,
	"dy": 128, "dz": 129,
	"ea": 130, "eb": 131, "ec": 132, "ed": 133, "ee": 134, "ef": 135, "eg": 136, "eh": 137,
	"ei": 138, "ej": 139, "ek": 140, "el": 141, "em": 142, "en": 143, "eo": 144, "ep": 145,
	"eq": 146, "er": 147, "es": 148, "et": 149, "eu": 150, "ev": 151, "ew": 152, "ex": 153,
	"ey": 154, "ez": 155,

	// 混合大小写
	"aA": 26, "aB": 27, "aC": 28, "aD": 29, "aE": 30, "aF": 31, "aG": 32, "aH": 33,
	"aI": 34, "aJ": 35, "aK": 36, "aL": 37, "aM": 38, "aN": 39, "aO": 40, "aP": 41,
	"aQ": 42, "aR": 43, "aS": 44, "aT": 45, "aU": 46, "aV": 47, "aW": 48, "aX": 49,
	"aY": 50, "aZ": 51,
	"bA": 52, "bB": 53, "bC": 54, "bD": 55, "bE": 56, "bF": 57, "bG": 58, "bH": 59,
	"bI": 60, "bJ": 61, "bK": 62, "bL": 63, "bM": 64, "bN": 65, "bO": 66, "bP": 67,
	"bQ": 68, "bR": 69, "bS": 70, "bT": 71, "bU": 72, "bV": 73, "bW": 74, "bX": 75,
	"bY": 76, "bZ": 77,
	"cA": 78, "cB": 79, "cC": 80, "cD": 81, "cE": 82, "cF": 83, "cG": 84, "cH": 85,
	"cI": 86, "cJ": 87, "cK": 88, "cL": 89, "cM": 90, "cN": 91, "cO": 92, "cP": 93,
	"cQ": 94, "cR": 95, "cS": 96, "cT": 97, "cU": 98, "cV": 99, "cW": 100, "cX": 101,
	"cY": 102, "cZ": 103,
	"dA": 104, "dB": 105, "dC": 106, "dD": 107, "dE": 108, "dF": 109, "dG": 110, "dH": 111,
	"dI": 112, "dJ": 113, "dK": 114, "dL": 115, "dM": 116, "dN": 117, "dO": 118, "dP": 119,
	"dQ": 120, "dR": 121, "dS": 122, "dT": 123, "dU": 124, "dV": 125, "dW": 126, "dX": 127,
	"dY": 128, "dZ": 129,
	"eA": 130, "eB": 131, "eC": 132, "eD": 133, "eE": 134, "eF": 135, "eG": 136, "eH": 137,
	"eI": 138, "eJ": 139, "eK": 140, "eL": 141, "eM": 142, "eN": 143, "eO": 144, "eP": 145,
	"eQ": 146, "eR": 147, "eS": 148, "eT": 149, "eU": 150, "eV": 151, "eW": 152, "eX": 153,
	"eY": 154, "eZ": 155,

	"Aa": 26, "Ab": 27, "Ac": 28, "Ad": 29, "Ae": 30, "Af": 31, "Ag": 32, "Ah": 33,
	"Ai": 34, "Aj": 35, "Ak": 36, "Al": 37, "Am": 38, "An": 39, "Ao": 40, "Ap": 41,
	"Aq": 42, "Ar": 43, "As": 44, "At": 45, "Au": 46, "Av": 47, "Aw": 48, "Ax": 49,
	"Ay": 50, "Az": 51,
	"Ba": 52, "Bb": 53, "Bc": 54, "Bd": 55, "Be": 56, "Bf": 57, "Bg": 58, "Bh": 59,
	"Bi": 60, "Bj": 61, "Bk": 62, "Bl": 63, "Bm": 64, "Bn": 65, "Bo": 66, "Bp": 67,
	"Bq": 68, "Br": 69, "Bs": 70, "Bt": 71, "Bu": 72, "Bv": 73, "Bw": 74, "Bx": 75,
	"By": 76, "Bz": 77,
	"Ca": 78, "Cb": 79, "Cc": 80, "Cd": 81, "Ce": 82, "Cf": 83, "Cg": 84, "Ch": 85,
	"Ci": 86, "Cj": 87, "Ck": 88, "Cl": 89, "Cm": 90, "Cn": 91, "Co": 92, "Cp": 93,
	"Cq": 94, "Cr": 95, "Cs": 96, "Ct": 97, "Cu": 98, "Cv": 99, "Cw": 100, "Cx": 101,
	"Cy": 102, "Cz": 103,
	"Da": 104, "Db": 105, "Dc": 106, "Dd": 107, "De": 108, "Df": 109, "Dg": 110, "Dh": 111,
	"Di": 112, "Dj": 113, "Dk": 114, "Dl": 115, "Dm": 116, "Dn": 117, "Do": 118, "Dp": 119,
	"Dq": 120, "Dr": 121, "Ds": 122, "Dt": 123, "Du": 124, "Dv": 125, "Dw": 126, "Dx": 127,
	"Dy": 128, "Dz": 129,
	"Ea": 130, "Eb": 131, "Ec": 132, "Ed": 133, "Ee": 134, "Ef": 135, "Eg": 136, "Eh": 137,
	"Ei": 138, "Ej": 139, "Ek": 140, "El": 141, "Em": 142, "En": 143, "Eo": 144, "Ep": 145,
	"Eq": 146, "Er": 147, "Es": 148, "Et": 149, "Eu": 150, "Ev": 151, "Ew": 152, "Ex": 153,
	"Ey": 154, "Ez": 155,
}
var ColumnNumberToNameCache = map[int]string{
	0: "A", 1: "B", 2: "C", 3: "D", 4: "E", 5: "F", 6: "G", 7: "H", 8: "I", 9: "J",
	10: "K", 11: "L", 12: "M", 13: "N", 14: "O", 15: "P", 16: "Q", 17: "R", 18: "S",
	19: "T", 20: "U", 21: "V", 22: "W", 23: "X", 24: "Y", 25: "Z",
	26: "AA", 27: "AB", 28: "AC", 29: "AD", 30: "AE", 31: "AF", 32: "AG", 33: "AH",
	34: "AI", 35: "AJ", 36: "AK", 37: "AL", 38: "AM", 39: "AN", 40: "AO", 41: "AP",
	42: "AQ", 43: "AR", 44: "AS", 45: "AT", 46: "AU", 47: "AV", 48: "AW", 49: "AX",
	50: "AY", 51: "AZ",
	52: "BA", 53: "BB", 54: "BC", 55: "BD", 56: "BE", 57: "BF", 58: "BG", 59: "BH",
	60: "BI", 61: "BJ", 62: "BK", 63: "BL", 64: "BM", 65: "BN", 66: "BO", 67: "BP",
	68: "BQ", 69: "BR", 70: "BS", 71: "BT", 72: "BU", 73: "BV", 74: "BW", 75: "BX",
	76: "BY", 77: "BZ",
	78: "CA", 79: "CB", 80: "CC", 81: "CD", 82: "CE", 83: "CF", 84: "CG", 85: "CH",
	86: "CI", 87: "CJ", 88: "CK", 89: "CL", 90: "CM", 91: "CN", 92: "CO", 93: "CP",
	94: "CQ", 95: "CR", 96: "CS", 97: "CT", 98: "CU", 99: "CV", 100: "CW", 101: "CX",
	102: "CY", 103: "CZ",
	104: "DA", 105: "DB", 106: "DC", 107: "DD", 108: "DE", 109: "DF", 110: "DG", 111: "DH",
	112: "DI", 113: "DJ", 114: "DK", 115: "DL", 116: "DM", 117: "DN", 118: "DO", 119: "DP",
	120: "DQ", 121: "DR", 122: "DS", 123: "DT", 124: "DU", 125: "DV", 126: "DW", 127: "DX",
	128: "DY", 129: "DZ",
	130: "EA", 131: "EB", 132: "EC", 133: "ED", 134: "EE", 135: "EF", 136: "EG", 137: "EH",
	138: "EI", 139: "EJ", 140: "EK", 141: "EL", 142: "EM", 143: "EN", 144: "EO", 145: "EP",
	146: "EQ", 147: "ER", 148: "ES", 149: "ET", 150: "EU", 151: "EV", 152: "EW", 153: "EX",
	154: "EY", 155: "EZ",
}

// index 从0开始
func ColumnNameToNumber(name string) (int, error) {
	if v, ok := ColumnNameToNumberCache[name]; ok {
		return v, nil
	}
	number, err := excelize.ColumnNameToNumber(name)
	if err == nil {
		ColumnNameToNumberCache[name] = number - 1
		return number - 1, nil
	}
	return 0, err
}

// index 从0开始
func ColumnNumberToName(num int) (string, error) {
	if v, ok := ColumnNumberToNameCache[num]; ok {
		return v, nil
	}
	name, err := excelize.ColumnNumberToName(num + 1)
	if err == nil {
		ColumnNumberToNameCache[num] = name
		return name, nil
	}
	return "", err
}
