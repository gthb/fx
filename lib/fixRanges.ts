import { isRange } from './isType.ts';
import { parseA1Ref, parseA1RefXlsx } from './parseA1Ref.ts';
import { stringifyA1Ref, stringifyA1RefXlsx } from './stringifyA1Ref.ts';
import { addA1RangeBounds } from './addA1RangeBounds.ts';
import { parseStructRef, parseStructRefXlsx } from './parseStructRef.ts';
import { stringifyStructRef, stringifyStructRefXlsx } from './stringifyStructRef.ts';
import { tokenize, type OptsTokenize, tokenizeXlsx } from './tokenize.ts';
import { REF_STRUCT } from './constants.ts';
import type { ReferenceA1, ReferenceA1Xlsx, Token } from './types.ts';
import { cloneToken } from './cloneToken.ts';
import { stringifyTokens } from './stringifyTokens.ts';
import { followsRangeOperator, quoteRefPrefix } from './stringifyPrefix.ts';

// There is no R1C1 counterpart to this. This is because without an anchor cell
// it is impossible to determine if a relative+absolute range (R[1]C[1]:R5C5)
// needs to be flipped or not. The solution is to convert to A1 first:
// translateToRC(fixRanges(translateToA1(...)))

// Why the sheet range of a 3-D reference is left alone (see the fixTokenRanges docstring): on
// entry, Excel normalizes a sheet range in three ways:
//
// - ordering the two sheet names ("Mar:Jan" to "Jan:Mar")
// - collapsing a degenerate range ("Jan:Jan" to "Jan")
// - correcting each name's case to the sheet's own
//
// All three need the workbook's sheet names, in order.

// Re-deciding a prefix's quotes from its sheet name alone loses the ones that are there for where
// the prefix stands rather than for what it is called: at a range operator's right endpoint, a
// bare prefix can merge with the operand on its left into a sheet range, so `Mar:'.Mar'!A1`
// written back bare would name the two sheets `Mar` and `.Mar` where the input named one. In that
// position a prefix therefore keeps the quotes it arrived with, on top of any its name calls for.
//
// It does not gain quotes it arrived without. Excel writes every prefix there quoted, and
// `translateToA1` follows it, but adding them here would rewrite references a caller never asked
// about; keeping them is what stops the reference changing, and it leaves `translateToA1`'s
// output alone rather than undoing it.
function keepPrefixQuotes (value: string, tokens: Token[], index: number): string {
  return (tokens[index].value.startsWith("'") && followsRangeOperator(tokens, index))
    ? quoteRefPrefix(value)
    : value;
}

/**
 * Options for {@link fixTokenRanges} and {@link fixFormulaRanges}.
 */
export type OptsFixRanges = {
  /**
   * Fill in any undefined bounds of range objects. Top to 0, bottom to 1048575, left to 0, and right to 16383.
   * @defaultValue false
   */
  addBounds?: boolean,
  /**
   * Enforces using the `[#This Row]` instead of the `@` shorthand when serializing structured ranges.
   * @defaultValue false
   */
  thisRow?: boolean,
};

/**
 * Normalizes A1 style ranges and structured references in a list of tokens.
 *
 * It ensures that that the top and left coordinates of an A1 range are on the
 * left-hand side of a colon operator:
 *
 * ```
 * B2:A1 → A1:B2
 * 1:A1 → A1:1
 * A:A1 → A1:A
 * B:A → A:B
 * 2:1 → 1:2
 * A1:A1 → A1
 * ```
 *
 * When `{ addBounds }` option is set to true, the missing bounds are also added.
 * This can be done to ensure Excel compatible ranges. The fixes then additionally include:
 *
 * ```
 * 1:A1 → A1:1 → 1:1
 * A:A1 → A1:A → A:A
 * A1:A → A:A
 * A1:1 → A:1
 * B2:B → B2:1048576
 * B2:2 → B2:XFD2
 * ```
 *
 * Structured ranges are normalized to have consistent order and capitalization
 * of sections as well as removing redundant ones.
 *
 * The sheet range of a 3-D reference is left alone, so `Sheet2:Sheet1!B2:A1`
 * becomes `Sheet2:Sheet1!A1:B2` and no further. Normalizing the prefix would
 * require the workbook's sheet names, in order, which _Fx_ does not know.
 *
 * Returns a new array of tokens with values and position data updated.
 *
 * @see {@link OptsFixRanges}
 * @param tokens A list of tokens to be adjusted.
 * @param [options]  Options.
 * @returns A token list with ranges adjusted.
 */
export function fixTokenRanges (
  tokens: Token[],
  options: OptsFixRanges = {}
): Token[] {
  if (!Array.isArray(tokens)) {
    throw new Error('fixRanges expects an array of tokens');
  }
  const { addBounds, thisRow } = options;
  let offsetSkew = 0;
  const output: Token[] = [];
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    const token = cloneToken(t);
    let offsetDelta = 0;
    if (token.type === REF_STRUCT) {
      const sref = parseStructRef(token.value);
      const newValue = stringifyStructRef(sref, { thisRow });
      offsetDelta = newValue.length - token.value.length;
      token.value = newValue;
    }
    else if (isRange(token)) {
      const ref = parseA1Ref(token.value, { allowTernary: true }) as ReferenceA1;
      const range = ref.range;
      // fill missing dimensions?
      if (addBounds) {
        addA1RangeBounds(range);
      }
      const newValue = keepPrefixQuotes(stringifyA1Ref(ref), tokens, i);
      offsetDelta = newValue.length - token.value.length;
      token.value = newValue;
    }
    // ensure that positioning is still correct
    if (offsetSkew || offsetDelta) {
      if (token.loc) {
        token.loc[0] += offsetSkew;
      }
      offsetSkew += offsetDelta;
      if (token.loc) {
        token.loc[1] += offsetSkew;
      }
    }
    else {
      offsetSkew += offsetDelta;
    }
    output.push(token);
  }

  return output;
}

/**
 * Normalizes A1 style ranges and structured references in a list of tokens.
 *
 * It ensures that that the top and left coordinates of an A1 range are on the
 * left-hand side of a colon operator:
 *
 * ```
 * B2:A1 → A1:B2
 * 1:A1 → A1:1
 * A:A1 → A1:A
 * B:A → A:B
 * 2:1 → 1:2
 * A1:A1 → A1
 * ```
 *
 * When `{ addBounds }` option is set to true, the missing bounds are also added.
 * This can be done to ensure Excel compatible ranges. The fixes then additionally include:
 *
 * ```
 * 1:A1 → A1:1 → 1:1
 * A:A1 → A1:A → A:A
 * A1:A → A:A
 * A1:1 → A:1
 * B2:B → B2:1048576
 * B2:2 → B2:XFD2
 * ```
 *
 * Structured ranges are normalized to have consistent order and capitalization
 * of sections as well as removing redundant ones.
 *
 * The sheet range of a 3-D reference is left alone, so `Sheet2:Sheet1!B2:A1`
 * becomes `Sheet2:Sheet1!A1:B2` and no further. Normalizing it would take the
 * workbook's sheet names, in order, which _Fx_ does not know.
 *
 * Returns a new array of tokens with values and position data updated.
 *
 * @see {@link OptsFixRanges}
 * @param tokens A list of tokens to be adjusted.
 * @param [options]  Options.
 * @returns A token list with ranges adjusted.
 */
export function fixTokenRangesXlsx (
  tokens: Token[],
  options: OptsFixRanges = {}
): Token[] {
  if (!Array.isArray(tokens)) {
    throw new Error('fixRanges expects an array of tokens');
  }
  const { addBounds, thisRow } = options;
  let offsetSkew = 0;
  const output: Token[] = [];
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    const token = cloneToken(t);
    let offsetDelta = 0;
    if (token.type === REF_STRUCT) {
      const sref = parseStructRefXlsx(token.value);
      const newValue = stringifyStructRefXlsx(sref, { thisRow });
      offsetDelta = newValue.length - token.value.length;
      token.value = newValue;
    }
    else if (isRange(token)) {
      const ref = parseA1RefXlsx(token.value, { allowTernary: true }) as ReferenceA1Xlsx;
      const range = ref.range;
      // fill missing dimensions?
      if (addBounds) {
        addA1RangeBounds(range);
      }
      const newValue = keepPrefixQuotes(stringifyA1RefXlsx(ref), tokens, i);
      offsetDelta = newValue.length - token.value.length;
      token.value = newValue;
    }
    // ensure that positioning is still correct
    if (offsetSkew || offsetDelta) {
      if (token.loc) {
        token.loc[0] += offsetSkew;
      }
      offsetSkew += offsetDelta;
      if (token.loc) {
        token.loc[1] += offsetSkew;
      }
    }
    else {
      offsetSkew += offsetDelta;
    }
    output.push(token);
  }

  return output;
}

/**
 * Normalizes A1 style ranges and structured references in a formula.
 *
 * Internally it uses {@link fixTokenRanges} so see it's documentation for details.
 *
 * Returns the same formula with the ranges updated. If an array of tokens was
 * supplied, then a new array is returned.
 *
 * @see {@link OptsFixRanges} & {@link OptsTokenize}
 * @param formula A string (an Excel formula) or a token list that should be adjusted.
 * @param [options]  Options
 * @returns A formula string with ranges adjusted
 */
export function fixFormulaRanges (
  formula: string,
  options: OptsFixRanges & OptsTokenize = {}
): string {
  if (typeof formula !== 'string') {
    throw new Error('fixFormulaRanges expects a string formula');
  }
  return stringifyTokens(
    fixTokenRanges(
      tokenize(formula, options),
      options
    )
  );
}

/**
 * Normalizes A1 style ranges and structured references in a formula.
 *
 * Internally it uses {@link fixTokenRanges} so see it's documentation for details.
 *
 * Returns the same formula with the ranges updated. If an array of tokens was
 * supplied, then a new array is returned.
 *
 * @see {@link OptsFixRanges} & {@link OptsTokenize}
 * @param formula A string (an Excel formula) or a token list that should be adjusted.
 * @param [options]  Options
 * @returns A formula string with ranges adjusted
 */
export function fixFormulaRangesXlsx (
  formula: string,
  options: OptsFixRanges & OptsTokenize = {}
): string {
  if (typeof formula !== 'string') {
    throw new Error('fixFormulaRanges expects a string formula');
  }
  return stringifyTokens(
    fixTokenRangesXlsx(
      tokenizeXlsx(formula, options),
      options
    )
  );
}
