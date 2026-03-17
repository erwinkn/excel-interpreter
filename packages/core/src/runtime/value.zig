pub const Value = union(enum) {
    blank,
    number: f64,
    boolean: bool,
};
