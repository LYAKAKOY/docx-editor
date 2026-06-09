/**
 * Line Spacing Picker Component (Radix UI)
 *
 * A dropdown selector for choosing line spacing values using Radix Select.
 * Styled like Google Docs with options: Single, 1.15, 1.5, Double
 */

import * as React from 'react';
import {
  Select,
  SelectContent,
  SelectGroup,
  SelectItem,
  SelectLabel,
  SelectSeparator,
  SelectTrigger,
} from './Select';
import { cn } from '../../lib/utils';
import { IconLineSpacing } from './Icons';
import { useTranslation } from '../../i18n';
import type { TranslationKey } from '@eigenpal/docx-editor-i18n';

// ============================================================================
// TYPES
// ============================================================================

export interface LineSpacingOption {
  label: string;
  labelKey?: TranslationKey;
  value: number;
  twipsValue: number;
}

export interface LineSpacingPickerProps {
  value?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  onChange?: (twipsValue: number) => void;
  onParagraphSpacingChange?: (side: 'before' | 'after', twipsValue: number) => void;
  options?: LineSpacingOption[];
  disabled?: boolean;
  className?: string;
  width?: number | string;
}

// ============================================================================
// CONSTANTS
// ============================================================================

/**
 * Standard line spacing options (Google Docs style)
 * OOXML uses twips for line spacing when lineRule="auto"
 * 240 twips = 1.0 line spacing (single)
 */
const DEFAULT_OPTIONS: LineSpacingOption[] = [
  { label: 'Single', labelKey: 'lineSpacing.single', value: 1.0, twipsValue: 240 },
  { label: '1.15', value: 1.15, twipsValue: 276 },
  { label: '1.5', value: 1.5, twipsValue: 360 },
  { label: 'Double', labelKey: 'lineSpacing.double', value: 2.0, twipsValue: 480 },
];

const DEFAULT_PARAGRAPH_SPACING_TWIPS = 240;
const TWIPS_PER_POINT = 20;

function formatParagraphSpacingPoints(twips: number | undefined): string {
  const points = (twips ?? 0) / TWIPS_PER_POINT;
  return Number.isInteger(points) ? String(points) : points.toFixed(2).replace(/\.?0+$/, '');
}

function parseParagraphSpacingPoints(value: string): number | null {
  const normalized = value.trim().replace(',', '.');
  if (!normalized) return 0;
  const points = Number(normalized);
  return Number.isFinite(points) && points >= 0 ? points : null;
}

function paragraphSpacingPointsToTwips(points: number): number {
  return Math.round(points * TWIPS_PER_POINT);
}

// ============================================================================
// COMPONENT
// ============================================================================

export function LineSpacingPicker({
  value,
  spaceBefore,
  spaceAfter,
  onChange,
  onParagraphSpacingChange,
  options = DEFAULT_OPTIONS,
  disabled = false,
  className,
}: LineSpacingPickerProps) {
  const { t } = useTranslation();
  const beforeInputId = React.useId();
  const afterInputId = React.useId();
  const [beforeInput, setBeforeInput] = React.useState(() =>
    formatParagraphSpacingPoints(spaceBefore)
  );
  const [afterInput, setAfterInput] = React.useState(() =>
    formatParagraphSpacingPoints(spaceAfter)
  );
  const skipNextBlurCommitRef = React.useRef<{ side: 'before' | 'after'; value: string } | null>(
    null
  );

  React.useEffect(() => {
    setBeforeInput(formatParagraphSpacingPoints(spaceBefore));
  }, [spaceBefore]);

  React.useEffect(() => {
    setAfterInput(formatParagraphSpacingPoints(spaceAfter));
  }, [spaceAfter]);

  // Find current option by twips value
  const currentOption = React.useMemo(() => {
    if (value === undefined) return options[0]; // Default to Single
    return options.find((opt) => opt.twipsValue === value) || options[0];
  }, [value, options]);

  const hasSpaceBefore = (spaceBefore ?? 0) > 0;
  const hasSpaceAfter = (spaceAfter ?? 0) > 0;

  const handleValueChange = React.useCallback(
    (newValue: string) => {
      if (newValue === 'paragraph-space-before') {
        onParagraphSpacingChange?.('before', hasSpaceBefore ? 0 : DEFAULT_PARAGRAPH_SPACING_TWIPS);
        return;
      }
      if (newValue === 'paragraph-space-after') {
        onParagraphSpacingChange?.('after', hasSpaceAfter ? 0 : DEFAULT_PARAGRAPH_SPACING_TWIPS);
        return;
      }

      const twips = parseInt(newValue, 10);
      if (!isNaN(twips)) {
        onChange?.(twips);
      }
    },
    [hasSpaceAfter, hasSpaceBefore, onChange, onParagraphSpacingChange]
  );

  const resetParagraphSpacingInput = React.useCallback(
    (side: 'before' | 'after') => {
      const formatted =
        side === 'before'
          ? formatParagraphSpacingPoints(spaceBefore)
          : formatParagraphSpacingPoints(spaceAfter);
      if (side === 'before') {
        setBeforeInput(formatted);
      } else {
        setAfterInput(formatted);
      }
    },
    [spaceAfter, spaceBefore]
  );

  const commitParagraphSpacingInput = React.useCallback(
    (side: 'before' | 'after', rawValue: string) => {
      const points = parseParagraphSpacingPoints(rawValue);
      if (points == null) {
        resetParagraphSpacingInput(side);
        return;
      }
      const twips = paragraphSpacingPointsToTwips(points);
      const formatted = formatParagraphSpacingPoints(twips);
      if (side === 'before') {
        setBeforeInput(formatted);
      } else {
        setAfterInput(formatted);
      }
      onParagraphSpacingChange?.(side, twips);
    },
    [onParagraphSpacingChange, resetParagraphSpacingInput]
  );

  const handleParagraphSpacingKeyDown = React.useCallback(
    (side: 'before' | 'after', event: React.KeyboardEvent<HTMLInputElement>) => {
      event.stopPropagation();
      if (event.key === 'Enter') {
        event.preventDefault();
        skipNextBlurCommitRef.current = { side, value: event.currentTarget.value };
        commitParagraphSpacingInput(side, event.currentTarget.value);
        event.currentTarget.blur();
      }
      if (event.key === 'Escape') {
        event.preventDefault();
        resetParagraphSpacingInput(side);
        event.currentTarget.blur();
      }
    },
    [commitParagraphSpacingInput, resetParagraphSpacingInput]
  );

  const handleParagraphSpacingBlur = React.useCallback(
    (side: 'before' | 'after', event: React.FocusEvent<HTMLInputElement>) => {
      const skipped = skipNextBlurCommitRef.current;
      if (skipped?.side === side && skipped.value === event.currentTarget.value) {
        skipNextBlurCommitRef.current = null;
        return;
      }
      if ((event.relatedTarget as HTMLElement | null)?.closest('[data-select-interactive]')) {
        return;
      }
      commitParagraphSpacingInput(side, event.currentTarget.value);
    },
    [commitParagraphSpacingInput]
  );

  const stopParagraphSpacingInteraction = React.useCallback((event: React.SyntheticEvent) => {
    event.stopPropagation();
  }, []);

  const getOptionLabel = (option: LineSpacingOption) =>
    option.labelKey ? t(option.labelKey) : option.label;

  return (
    <Select
      value={currentOption.twipsValue.toString()}
      onValueChange={handleValueChange}
      disabled={disabled}
    >
      <SelectTrigger
        className={cn('h-8 text-sm gap-0.5 px-2', className)}
        style={{ width: 'auto' }}
        aria-label={t('lineSpacing.label')}
        title={t('lineSpacing.lineSpacingTitle', { label: getOptionLabel(currentOption) })}
      >
        <IconLineSpacing className="h-5 w-5 shrink-0" />
      </SelectTrigger>
      <SelectContent>
        {options.map((option) => (
          <SelectItem key={option.twipsValue} value={option.twipsValue.toString()}>
            {getOptionLabel(option)}
          </SelectItem>
        ))}
        <SelectSeparator />
        <SelectGroup>
          <SelectLabel>{t('lineSpacing.paragraphSpacing')}</SelectLabel>
          <div
            data-select-interactive
            className="px-2 pb-1.5 pt-1"
            onClick={stopParagraphSpacingInteraction}
            onMouseDown={stopParagraphSpacingInteraction}
            onPointerDown={stopParagraphSpacingInteraction}
          >
            <div className="grid grid-cols-[minmax(3.5rem,auto)_4.5rem_auto] items-center gap-2 text-sm text-slate-700">
              <label htmlFor={beforeInputId} className="whitespace-nowrap text-xs text-slate-600">
                {t('lineSpacing.spaceBeforePt')}
              </label>
              <input
                id={beforeInputId}
                aria-label={t('lineSpacing.spaceBeforePtAria')}
                className="h-7 rounded border border-slate-200 px-2 text-sm text-slate-900 outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500"
                type="number"
                min={0}
                step={0.5}
                inputMode="decimal"
                value={beforeInput}
                onChange={(event) => setBeforeInput(event.currentTarget.value)}
                onBlur={(event) => handleParagraphSpacingBlur('before', event)}
                onKeyDown={(event) => handleParagraphSpacingKeyDown('before', event)}
              />
              <span className="text-xs text-slate-500">{t('lineSpacing.pointsAbbrev')}</span>
              <label htmlFor={afterInputId} className="whitespace-nowrap text-xs text-slate-600">
                {t('lineSpacing.spaceAfterPt')}
              </label>
              <input
                id={afterInputId}
                aria-label={t('lineSpacing.spaceAfterPtAria')}
                className="h-7 rounded border border-slate-200 px-2 text-sm text-slate-900 outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500"
                type="number"
                min={0}
                step={0.5}
                inputMode="decimal"
                value={afterInput}
                onChange={(event) => setAfterInput(event.currentTarget.value)}
                onBlur={(event) => handleParagraphSpacingBlur('after', event)}
                onKeyDown={(event) => handleParagraphSpacingKeyDown('after', event)}
              />
              <span className="text-xs text-slate-500">{t('lineSpacing.pointsAbbrev')}</span>
            </div>
          </div>
          <SelectItem value="paragraph-space-before">
            {t(hasSpaceBefore ? 'lineSpacing.removeSpaceBefore' : 'lineSpacing.addSpaceBefore')}
          </SelectItem>
          <SelectItem value="paragraph-space-after">
            {t(hasSpaceAfter ? 'lineSpacing.removeSpaceAfter' : 'lineSpacing.addSpaceAfter')}
          </SelectItem>
        </SelectGroup>
      </SelectContent>
    </Select>
  );
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

export function getDefaultLineSpacingOptions(): LineSpacingOption[] {
  return [...DEFAULT_OPTIONS];
}

export function lineSpacingMultiplierToTwips(multiplier: number): number {
  return Math.round(multiplier * 240);
}

export function twipsToLineSpacingMultiplier(twips: number): number {
  return twips / 240;
}

export function getLineSpacingLabel(twips: number): string {
  const option = DEFAULT_OPTIONS.find((opt) => opt.twipsValue === twips);
  if (option) return option.label;
  const multiplier = twipsToLineSpacingMultiplier(twips);
  return multiplier.toFixed(2).replace(/\.?0+$/, '');
}

export function isStandardLineSpacing(twips: number): boolean {
  return DEFAULT_OPTIONS.some((opt) => opt.twipsValue === twips);
}

export function nearestStandardLineSpacing(twips: number): LineSpacingOption {
  let nearest = DEFAULT_OPTIONS[0];
  let minDiff = Math.abs(twips - nearest.twipsValue);

  for (const option of DEFAULT_OPTIONS) {
    const diff = Math.abs(twips - option.twipsValue);
    if (diff < minDiff) {
      minDiff = diff;
      nearest = option;
    }
  }

  return nearest;
}

export function createLineSpacingOption(multiplier: number): LineSpacingOption {
  const twipsValue = lineSpacingMultiplierToTwips(multiplier);
  const label = multiplier.toFixed(2).replace(/\.?0+$/, '');
  return { label, value: multiplier, twipsValue };
}

export default LineSpacingPicker;
