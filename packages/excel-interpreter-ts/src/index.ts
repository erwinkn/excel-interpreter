export type ExcelScalar = string | number | boolean | null;

export interface ExcelContext {
  getCell(address: string, sheet?: string): ExcelScalar;
}

export class ExpressionBuilder {
  private parts: string[] = [];

  ref(address: string, sheet?: string): this {
    this.parts.push(sheet ? `${sheet}!${address}` : address);
    return this;
  }

  lit(value: ExcelScalar): this {
    if (value === null) {
      this.parts.push("\"\"");
    } else if (typeof value === "string") {
      this.parts.push(`\"${value.replaceAll('"', '""')}\"`);
    } else if (typeof value === "boolean") {
      this.parts.push(value ? "TRUE" : "FALSE");
    } else {
      this.parts.push(String(value));
    }
    return this;
  }

  fn(name: string, ...args: string[]): this {
    this.parts.push(`${name.toUpperCase()}(${args.join(",")})`);
    return this;
  }

  raw(fragment: string): this {
    this.parts.push(fragment);
    return this;
  }

  build(): string {
    return this.parts.join("");
  }
}

export function evaluateFormula(formula: string, _context: ExcelContext): ExcelScalar {
  // Placeholder evaluator. Real workbook-aware evaluation will be added in later iterations.
  if (!formula.startsWith("=")) {
    return formula;
  }
  throw new Error("evaluateFormula is not implemented yet.");
}
