pub enum Lock<T> {
    Locked(T),
    Unlocked(T),
}

impl<T> Lock<T> where T: std::string::ToString {
    fn render(&self) -> String {
        match self {
            Lock::Locked(v) => format!("${}", v.to_string()),
            Lock::Unlocked(v) => v.to_string()
        }
    }
}

pub struct CellReference {
    pub column: Lock<String>,
    pub row: Option<Lock<i32>>,
}

impl CellReference {
    fn render(&self) -> String {
        format!("{}{}", self.column.render(), self.row.as_ref().map(|x| x.render()).unwrap_or_else(|| String::new()))
    }
}

impl ValueReference for CellReference {
    fn render(&self) -> String {
        self.render()
    }
}

pub struct CellRange {
    pub sheet: Option<String>,
    pub from: CellReference,
    pub to: Option<CellReference>,
}

impl CellRange {
    fn render(&self) -> String {
        let prefix = if let Some(sheet) = self.sheet.as_ref() {
            format!("'{}'!", sheet)
        } else {
            String::new()
        };

        let xx = if let Some(to) = self.to.as_ref() {
            format!(":{}", to.render())
        } else {
            String::new()
        };

        format!("{}{}{}", prefix, self.from.render(), xx)
    }
}

impl ValueReference for CellRange {
    fn render(&self) -> String {
        self.render()
    }
}

#[macro_export]
macro_rules! blank {
    ($x:expr) => {{
        CellValue::Blank($x)
    }};
    () => {{
        CellValue::Blank(1)
    }};
}

#[macro_export]
macro_rules! cell {
    ($c:expr,$r:expr) => {
        CellReference {
            column: $c,
            row: Some($r),
        }
    };
    ($x:expr) => {
        CellReference {
            column: $x,
            row: None,
        }
    };
}

#[macro_export]
macro_rules! locked {
    ($x:expr) => {
        Lock::Locked($x)
    };
}

#[macro_export]
macro_rules! unlocked {
    ($x:expr) => {
        Lock::Unlocked($x)
    };
}

pub trait ValueReference {
    fn render(&self) -> String;
}

pub enum Formula {
    Concat(Vec<Box<ValueReference>>),
    IfError(Box<Formula>, Box<ValueReference>),
    VerticalLookUp(Box<ValueReference>, CellRange, i32, bool),
    Sum(Vec<Box<ValueReference>>),
    Subtract(Box<ValueReference>, Box<ValueReference>),
    Divide(Box<ValueReference>, Box<ValueReference>),
    Product(Vec<Box<ValueReference>>),
    Round(Box<ValueReference>, i32),
}

impl ValueReference for String {
    fn render(&self) -> String {
        format!("\"{}\"", self)
    }
}
impl ValueReference for &str {
    fn render(&self) -> String {
        format!("\"{}\"", self)
    }
}
impl ValueReference for f64 {
    fn render(&self) -> String {
        self.to_string()
    }
}

impl ValueReference for Formula {
    fn render(&self) -> String {
        render_formula(self)
    }
}

pub(crate) fn render_formula(formula: &Formula) -> String {
    match *formula {
        Formula::Concat(ref v) => format!("CONCATENATE({})", v.iter().map(|x| x.render()).collect::<Vec<String>>().join(",")),
        Formula::IfError(ref f, ref v) => format!("IFERROR({},{})", render_formula(f), v.render()),
        Formula::VerticalLookUp(ref a, ref b, ref c, ref d) => format!("VLOOKUP({},{},{},{})", a.render(), b.render(), c, d.to_string().to_uppercase()),
        Formula::Sum(ref v) => format!("SUM({})", v.iter().map(|x| x.render()).collect::<Vec<String>>().join(",")),
        Formula::Subtract(ref v1, ref v2) => format!("{}-{}", v1.render(), v2.render()),
        Formula::Divide(ref v1, ref v2) => format!("{}/{}", v1.render(), v2.render()),
        Formula::Product(ref v) => format!("PRODUCT({})", v.iter().map(|x| x.render()).collect::<Vec<String>>().join(",")),
        Formula::Round(ref v, d) => format!("ROUND({},{})", v.render(), d),
    }
}

mod tests {
    #[test]
    fn test_formula() {
        use crate::formula::{Formula::*, *};

        let x1 = CellReference { column: locked!("A".to_string()), row: Some(unlocked!(2)) };
        let x3 = CellReference { column: unlocked!("B".to_string()), row: Some(locked!(1)) };
        let concat = Concat(vec![Box::new(x1), Box::new("/"), Box::new(x3)]);
        let range = CellRange {
            sheet: Some("Data".to_string()),
            from: cell!(locked!("C".to_string()), locked!(1)),
            to: Some(cell!(locked!("D".to_string()), unlocked!(542))),
        };
        let f = IfError(Box::new(VerticalLookUp(Box::new(concat), range, 2, false)), Box::new("0"));
        let rendered = render_formula(&f);
        assert!(rendered == "IFERROR(VLOOKUP(CONCATENATE($A2,\"/\",B$1),'Data'!$C$1:$D542,2,FALSE),\"0\")")
    }
}
