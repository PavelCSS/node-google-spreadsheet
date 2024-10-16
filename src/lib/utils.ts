import * as _ from './lodash';
import {A1Range, GridRange, GridRangeWithoutWorksheetId, NamedRange} from './types/sheets-types';

export function getFieldMask(obj: Record<string, unknown>) {
  let fromGrid = '';
  const fromRoot = Object.keys(obj).filter((key) => key !== 'gridProperties').join(',');

  if (obj.gridProperties) {
    fromGrid = Object.keys(obj.gridProperties).map((key) => `gridProperties.${key}`).join(',');
    if (fromGrid.length && fromRoot.length) {
      fromGrid = `${fromGrid},`;
    }
  }
  return fromGrid + fromRoot;
}

export function columnToLetter(column: number) {
  let temp;
  let letter = '';
  let col = column;
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

export function letterToColumn(letter: string) {
  let column = 0;
  const { length } = letter;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * 26 ** (length - i - 1);
  }
  return column;
}

// send arrays in params with duplicate keys - ie `?thing=1&thing=2` vs `?thing[]=1...`
// solution taken from https://github.com/axios/axios/issues/604
export function axiosParamsSerializer(params: Record<PropertyKey, any>) {
  let options = '';
  Object.keys(params).forEach((key) => {
    const isParamTypeObject = typeof params[key] === 'object';
    const isParamTypeArray = isParamTypeObject && (params[key].length >= 0);
    if (!isParamTypeObject) options += `${key}=${encodeURIComponent(params[key])}&`;
    if (isParamTypeObject && isParamTypeArray) {
      // eslint-disable-next-line no-restricted-syntax
      for (const val of params[key]) {
        options += `${key}=${encodeURIComponent(val)}&`;
      }
    }
  });
  return options ? options.slice(0, -1) : options;
}


export function checkForDuplicateHeaders(headers: string[]) {
  // check for duplicate headers
  const checkForDupes = _.groupBy(headers); // { c1: ['c1'], c2: ['c2', 'c2' ]}
  _.each(checkForDupes, (grouped, header) => {
    if (!header) return; // empty columns are skipped, so multiple is ok
    if (grouped.length > 1) {
      throw new Error(`Duplicate header detected: "${header}". Please make sure all non-empty headers are unique`);
    }
  });
}

export const parseRangeA1 = (a1Address: A1Range) => {
  const split = a1Address.match(/([A-Z]+)?(\d+)?:?([A-Z]+)?(\d+)?$/) || [];
  if (!split) throw new Error(`Range address "${a1Address}" not valid`);
  const [rangeA1 = '', startColumnA1, startRowA1, endColumnA1, endRowA1] = split;

  const startColumnIndex = startColumnA1 ? letterToColumn(startColumnA1) - 1 : 0;
  const endColumnIndex = endColumnA1 ? letterToColumn(endColumnA1) : undefined;
  const startRowIndex = startRowA1 ? parseInt(startRowA1) - 1 : 0;
  const endRowIndex = endRowA1 ? parseInt(endRowA1) : undefined;

  return {
    rangeA1,
    startColumnIndex,
    endColumnIndex,
    startRowIndex,
    endRowIndex,
    startColumnA1,
    endColumnA1,
    startRowA1,
    endRowA1,
  };
};

export const toA1Range = ({
  startColumnIndex,
  endColumnIndex,
  startRowIndex,
  endRowIndex,
}: GridRangeWithoutWorksheetId) => {
  let startA1Range = '';
  let endA1Range = '';

  if (Number.isInteger(startColumnIndex)) {
    startA1Range += columnToLetter(startColumnIndex! + 1);
  }
  if (Number.isInteger(startRowIndex)) {
    startA1Range += `${startRowIndex! + 1}`;
  }

  if (Number.isInteger(endColumnIndex)) {
    endA1Range += columnToLetter(endColumnIndex!);
  }
  if (Number.isInteger(endRowIndex)) {
    endA1Range += `${endRowIndex!}`;
  }

  if ((startA1Range === endA1Range) || (startA1Range && !endA1Range)) {
    return startA1Range;
  }

  return [startA1Range, endA1Range].join(':');
};
