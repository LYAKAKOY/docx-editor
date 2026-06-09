<!--
  Mirrors packages/react/src/components/ui/HorizontalRuler.tsx 1:1.
  Same constants, same eighth-inch tick generation (heights 10/6/4/2 px),
  same CSS-triangle indent handles, same margin-zone backgrounds.
-->
<template>
  <div
    ref="rulerRef"
    class="docx-horizontal-ruler"
    :style="containerStyle"
    role="slider"
    aria-label="Horizontal ruler"
    :aria-valuemin="0"
    :aria-valuemax="pageWidthTwips"
  >
    <!-- Margin zones (drag to adjust margin) -->
    <div
      class="docx-horizontal-ruler__margin"
      :style="leftMarginStyle"
      @mousedown.prevent.stop="editable ? startDrag('leftMargin', $event) : null"
    />
    <div
      class="docx-horizontal-ruler__margin"
      :style="rightMarginStyle"
      @mousedown.prevent.stop="editable ? startDrag('rightMargin', $event) : null"
    />

    <!-- Tick marks -->
    <div class="docx-horizontal-ruler__ticks">
      <template v-for="(tick, i) in ticks" :key="i">
        <div
          class="docx-horizontal-ruler__tick-line"
          :style="{ left: tick.position + 'px', height: tick.height + 'px' }"
        />
        <div
          v-if="tick.label"
          class="docx-horizontal-ruler__tick-label"
          :style="{ left: tick.position + 'px' }"
        >
          {{ tick.label }}
        </div>
      </template>
    </div>

    <!-- Table column boundaries: Word-style gray handles shown while inside a table -->
    <div
      v-for="(boundary, i) in tableBoundaries ?? []"
      :key="`${boundary.tablePmStart}-${boundary.kind}-${boundary.columnIndex}-${i}`"
      class="docx-ruler-table-boundary"
      :style="tableBoundaryStyle(boundary, i)"
      role="slider"
      :aria-label="t('ruler.tableColumnBoundary')"
      aria-orientation="horizontal"
      :tabindex="editable ? 0 : -1"
      @mousedown.prevent.stop="startTableBoundaryDrag(boundary, $event)"
      @mouseenter="hoveredTableBoundary = i"
      @mouseleave="hoveredTableBoundary = null"
    >
      <div class="docx-ruler-table-boundary__line" />
    </div>

    <!-- First-line indent: ▼ down at top -->
    <div
      v-if="showFirstLineIndent"
      class="docx-ruler-indent"
      :style="indentContainerStyle('down', firstLinePosPx, dragging === 'firstLineIndent')"
      role="slider"
      :aria-label="t('ruler.firstLineIndent')"
      aria-orientation="horizontal"
      :tabindex="editable ? 0 : -1"
      @mousedown.prevent.stop="editable ? startDrag('firstLineIndent', $event) : null"
      @mouseenter="hovered = 'firstLineIndent'"
      @mouseleave="hovered = null"
    >
      <div :style="triangleStyle('down', triColor('firstLineIndent'))" />
    </div>

    <!-- Left indent: ▲ up at bottom -->
    <div
      v-if="editable"
      class="docx-ruler-indent"
      :style="indentContainerStyle('up', leftIndentPosPx, dragging === 'leftIndent')"
      role="slider"
      :aria-label="t('ruler.leftIndent')"
      aria-orientation="horizontal"
      :tabindex="editable ? 0 : -1"
      @mousedown.prevent.stop="startDrag('leftIndent', $event)"
      @mouseenter="hovered = 'leftIndent'"
      @mouseleave="hovered = null"
    >
      <div :style="triangleStyle('up', triColor('leftIndent'))" />
    </div>

    <!-- Right indent: ▲ up at bottom -->
    <div
      v-if="editable"
      class="docx-ruler-indent"
      :style="indentContainerStyle('up', rightIndentPosPx, dragging === 'rightIndent')"
      role="slider"
      :aria-label="t('ruler.rightIndent')"
      aria-orientation="horizontal"
      :tabindex="editable ? 0 : -1"
      @mousedown.prevent.stop="startDrag('rightIndent', $event)"
      @mouseenter="hovered = 'rightIndent'"
      @mouseleave="hovered = null"
    >
      <div :style="triangleStyle('up', triColor('rightIndent'))" />
    </div>

    <!-- Tab stops -->
    <div
      v-for="(tab, i) in tabStopPositions"
      :key="i"
      class="docx-horizontal-ruler__tab"
      :style="{ left: tab.px + 'px' }"
      :title="`${tab.label}`"
      @dblclick.prevent="$emit('tab-stop-remove', tab.twips)"
    >
      L
    </div>

    <!-- Drag tooltip -->
    <div
      v-if="dragging && tooltipText"
      class="docx-horizontal-ruler__tooltip"
      :style="{ left: tooltipX + 'px' }"
    >
      {{ tooltipText }}
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, onBeforeUnmount, type CSSProperties } from 'vue';
import type { TableRulerCellArea } from '@eigenpal/docx-editor-core/layout-bridge';
import type { TableRulerBoundary } from '@eigenpal/docx-editor-core/prosemirror/tableResize';
import type { SectionProperties, TabStop } from '@eigenpal/docx-editor-core/types/document';
import { twipsToPixels, pixelsToTwips } from '@eigenpal/docx-editor-core/utils/units';
import { useTranslation } from '../../i18n';

type MarkerType =
  | 'leftMargin'
  | 'rightMargin'
  | 'firstLineIndent'
  | 'leftIndent'
  | 'rightIndent'
  | 'tableBoundary';

const props = withDefaults(
  defineProps<{
    sectionProps?: SectionProperties | null;
    zoom?: number;
    editable?: boolean;
    showFirstLineIndent?: boolean;
    firstLineIndent?: number;
    hangingIndent?: boolean;
    indentLeft?: number;
    indentRight?: number;
    unit?: 'inch' | 'cm';
    tabStops?: TabStop[] | null;
    indentArea?: TableRulerCellArea | null;
    tableBoundaries?: TableRulerBoundary[] | null;
  }>(),
  {
    zoom: 1,
    editable: true,
    showFirstLineIndent: true,
    firstLineIndent: 0,
    hangingIndent: false,
    indentLeft: 0,
    indentRight: 0,
    unit: 'inch',
  }
);

const emit = defineEmits<{
  (e: 'left-margin-change', twips: number): void;
  (e: 'right-margin-change', twips: number): void;
  (e: 'first-line-indent-change', twips: number): void;
  (e: 'indent-left-change', twips: number): void;
  (e: 'indent-right-change', twips: number): void;
  (e: 'tab-stop-remove', twips: number): void;
  (e: 'table-boundary-change', boundary: TableRulerBoundary, twips: number): void;
}>();

const { t } = useTranslation();

// Mirror React HorizontalRuler.tsx:50-63
const DEFAULT_PAGE_WIDTH_TWIPS = 12240;
const DEFAULT_MARGIN_TWIPS = 1440;
const TWIPS_PER_INCH = 1440;
const TWIPS_PER_CM = 567;
const RULER_HEIGHT = 22;
const INDENT_COLOR = '#4285f4';
const INDENT_HOVER_COLOR = '#3367d6';
const INDENT_ACTIVE_COLOR = '#2a56c6';
const TABLE_MARKER_COLOR = '#f1f3f4';
const TABLE_MARKER_BORDER = '#b9c0c8';
const TABLE_MARKER_ACTIVE = '#e3e8ef';
const TRI_SIZE = 5;
const TRI_HEIGHT = Math.round(TRI_SIZE * 1.6); // 8

const rulerRef = ref<HTMLElement | null>(null);
const dragging = ref<MarkerType | null>(null);
const hovered = ref<MarkerType | null>(null);
const hoveredTableBoundary = ref<number | null>(null);
const activeTableBoundary = ref<TableRulerBoundary | null>(null);
const tooltipX = ref(0);
const tooltipText = ref('');

// Wrappers around core's twipsToPixels/pixelsToTwips that fold in the
// current zoom factor (core helpers are zoom-agnostic).
function tw2px(twips: number): number {
  return twipsToPixels(twips) * props.zoom;
}
function px2tw(px: number): number {
  return Math.round(pixelsToTwips(px / props.zoom));
}

const pageWidthTwips = computed(() => props.sectionProps?.pageWidth ?? DEFAULT_PAGE_WIDTH_TWIPS);
const leftMarginTwips = computed(() => props.sectionProps?.marginLeft ?? DEFAULT_MARGIN_TWIPS);
const rightMarginTwips = computed(() => props.sectionProps?.marginRight ?? DEFAULT_MARGIN_TWIPS);
const indentAreaLeftTwips = computed(() => props.indentArea?.leftTwips ?? leftMarginTwips.value);
const indentAreaRightTwips = computed(
  () => props.indentArea?.rightTwips ?? pageWidthTwips.value - rightMarginTwips.value
);
const indentAreaWidthTwips = computed(() =>
  Math.max(0, indentAreaRightTwips.value - indentAreaLeftTwips.value)
);

const pageWidthPx = computed(() => tw2px(pageWidthTwips.value));
const leftMarginPx = computed(() => tw2px(leftMarginTwips.value));
const rightMarginPx = computed(() => tw2px(rightMarginTwips.value));
const indentAreaLeftPx = computed(() => tw2px(indentAreaLeftTwips.value));
const indentAreaRightPx = computed(() => tw2px(indentAreaRightTwips.value));
const indentLeftPx = computed(() => tw2px(props.indentLeft));
const indentRightPx = computed(() => tw2px(props.indentRight));
const effectiveFirstLine = computed(() =>
  props.hangingIndent ? -Math.abs(props.firstLineIndent) : props.firstLineIndent
);
const firstLineIndentPx = computed(() => tw2px(effectiveFirstLine.value));

const leftIndentPosPx = computed(() => indentAreaLeftPx.value + indentLeftPx.value);
const rightIndentPosPx = computed(() => indentAreaRightPx.value - indentRightPx.value);
const firstLinePosPx = computed(
  () => indentAreaLeftPx.value + indentLeftPx.value + firstLineIndentPx.value
);

const containerStyle = computed<CSSProperties>(() => ({
  position: 'relative',
  width: pageWidthPx.value + 'px',
  height: RULER_HEIGHT + 'px',
  backgroundColor: 'transparent',
  overflow: 'visible',
  userSelect: 'none',
  cursor: dragging.value ? 'ew-resize' : 'default',
}));

const leftMarginStyle = computed<CSSProperties>(() => ({
  position: 'absolute',
  top: 0,
  left: 0,
  width: leftMarginPx.value + 'px',
  height: RULER_HEIGHT + 'px',
  backgroundColor: 'rgba(0, 0, 0, 0.02)',
  borderRight: '1px solid rgba(0,0,0,0.06)',
  cursor: props.editable ? 'ew-resize' : 'default',
  zIndex: 1,
}));
const rightMarginStyle = computed<CSSProperties>(() => ({
  position: 'absolute',
  top: 0,
  right: 0,
  width: rightMarginPx.value + 'px',
  height: RULER_HEIGHT + 'px',
  backgroundColor: 'rgba(0, 0, 0, 0.02)',
  borderLeft: '1px solid rgba(0,0,0,0.06)',
  cursor: props.editable ? 'ew-resize' : 'default',
  zIndex: 1,
}));

// Mirror generateTicks() in React HorizontalRuler.tsx:549.
// Eighth-inch ticks; major every inch (height 10), quarters height 6,
// halves height 4, eighths height 2.
const ticks = computed(() => {
  const out: { position: number; height: number; label?: string }[] = [];
  if (props.unit === 'inch') {
    const eighth = TWIPS_PER_INCH / 8;
    const total = Math.ceil(pageWidthTwips.value / eighth);
    for (let i = 0; i <= total; i++) {
      const twPos = i * eighth;
      if (twPos > pageWidthTwips.value) break;
      const px = tw2px(twPos);
      if (i % 8 === 0) {
        out.push({ position: px, height: 10, label: i / 8 > 0 ? String(i / 8) : undefined });
      } else if (i % 4 === 0) {
        out.push({ position: px, height: 6 });
      } else if (i % 2 === 0) {
        out.push({ position: px, height: 4 });
      } else {
        out.push({ position: px, height: 2 });
      }
    }
  } else {
    const mm = TWIPS_PER_CM / 10;
    const total = Math.ceil(pageWidthTwips.value / mm);
    for (let i = 0; i <= total; i++) {
      const twPos = i * mm;
      if (twPos > pageWidthTwips.value) break;
      const px = tw2px(twPos);
      if (i % 10 === 0) {
        out.push({ position: px, height: 10, label: i / 10 > 0 ? String(i / 10) : undefined });
      } else if (i % 5 === 0) {
        out.push({ position: px, height: 6 });
      } else {
        out.push({ position: px, height: 3 });
      }
    }
  }
  return out;
});

const tabStopPositions = computed(() => {
  if (!props.tabStops?.length) return [];
  return props.tabStops.map((ts) => {
    const pos = (ts as any).position ?? ts.pos ?? 0;
    return { px: tw2px(pos), twips: pos, label: formatValue(pos) };
  });
});

function formatValue(twips: number): string {
  if (props.unit === 'cm') return (twips / TWIPS_PER_CM).toFixed(1) + ' cm';
  return (twips / TWIPS_PER_INCH).toFixed(2) + '"';
}

function clampRounded(value: number, min: number, max: number): number {
  return Math.round(Math.max(min, Math.min(value, max)));
}

function triColor(marker: MarkerType): string {
  if (dragging.value === marker) return INDENT_ACTIVE_COLOR;
  if (hovered.value === marker) return INDENT_HOVER_COLOR;
  return INDENT_COLOR;
}

function indentContainerStyle(
  direction: 'up' | 'down',
  positionPx: number,
  isDragging: boolean
): CSSProperties {
  return {
    position: 'absolute',
    left: positionPx - TRI_SIZE + 'px',
    width: TRI_SIZE * 2 + 'px',
    height: TRI_HEIGHT + 2 + 'px',
    cursor: props.editable ? 'ew-resize' : 'default',
    zIndex: isDragging ? 10 : 4,
    ...(direction === 'down' ? { top: 0 } : { bottom: 0 }),
  };
}

function triangleStyle(direction: 'up' | 'down', color: string): CSSProperties {
  if (direction === 'down') {
    return {
      position: 'absolute',
      top: '1px',
      left: 0,
      width: 0,
      height: 0,
      borderLeft: `${TRI_SIZE}px solid transparent`,
      borderRight: `${TRI_SIZE}px solid transparent`,
      borderTop: `${TRI_HEIGHT}px solid ${color}`,
      transition: 'border-top-color 0.1s',
    };
  }
  return {
    position: 'absolute',
    bottom: '1px',
    left: 0,
    width: 0,
    height: 0,
    borderLeft: `${TRI_SIZE}px solid transparent`,
    borderRight: `${TRI_SIZE}px solid transparent`,
    borderBottom: `${TRI_HEIGHT}px solid ${color}`,
    transition: 'border-bottom-color 0.1s',
  };
}

function sameBoundary(a: TableRulerBoundary | null, b: TableRulerBoundary): boolean {
  return (
    !!a && a.tablePmStart === b.tablePmStart && a.columnIndex === b.columnIndex && a.kind === b.kind
  );
}

function tableBoundaryStyle(boundary: TableRulerBoundary, index: number): CSSProperties {
  const isDragging =
    dragging.value === 'tableBoundary' && sameBoundary(activeTableBoundary.value, boundary);
  const isHovered = hoveredTableBoundary.value === index;
  return {
    position: 'absolute',
    left: tw2px(boundary.positionTwips) - 8 + 'px',
    top: '1px',
    width: '16px',
    height: RULER_HEIGHT - 3 + 'px',
    border: `1px solid ${TABLE_MARKER_BORDER}`,
    backgroundColor: isDragging || isHovered ? TABLE_MARKER_ACTIVE : TABLE_MARKER_COLOR,
    boxSizing: 'border-box',
    cursor: props.editable ? 'col-resize' : 'default',
    zIndex: isDragging ? 11 : 5,
  };
}

function initialDragValue(type: MarkerType): number {
  if (type === 'leftMargin') return leftMarginTwips.value;
  if (type === 'rightMargin') return rightMarginTwips.value;
  if (type === 'leftIndent') return props.indentLeft;
  if (type === 'rightIndent') return props.indentRight;
  if (type === 'firstLineIndent') return props.firstLineIndent;
  return activeTableBoundary.value?.positionTwips ?? 0;
}

function beginDrag(event: MouseEvent, initialValue: number) {
  tooltipX.value = event.clientX - (rulerRef.value?.getBoundingClientRect().left ?? 0);
  tooltipText.value = formatValue(initialValue);

  document.addEventListener('mousemove', handleMove);
  document.addEventListener('mouseup', handleUp);
}

function startDrag(type: MarkerType, event: MouseEvent) {
  if (!props.editable) return;
  dragging.value = type;
  beginDrag(event, initialDragValue(type));
}

function startTableBoundaryDrag(boundary: TableRulerBoundary, event: MouseEvent) {
  if (!props.editable) return;
  dragging.value = 'tableBoundary';
  activeTableBoundary.value = boundary;
  beginDrag(event, boundary.positionTwips);
}

function handleMove(e: MouseEvent) {
  if (!dragging.value || !rulerRef.value) return;
  const rect = rulerRef.value.getBoundingClientRect();
  if (!rect) return;
  const x = e.clientX - rect.left;
  const positionTwips = px2tw(x);
  const value = computeNewValue(dragging.value, positionTwips);
  tooltipX.value = dragging.value === 'tableBoundary' ? tw2px(value) : x;
  tooltipText.value = formatValue(value);
  emitChange(dragging.value, value);
}

function handleUp(_e: MouseEvent) {
  dragging.value = null;
  activeTableBoundary.value = null;
  document.removeEventListener('mousemove', handleMove);
  document.removeEventListener('mouseup', handleUp);
}

function computeNewValue(marker: MarkerType, positionTwips: number): number {
  if (marker === 'tableBoundary') {
    const boundary = activeTableBoundary.value;
    if (!boundary) return Math.round(positionTwips);
    return clampRounded(positionTwips, boundary.minTwips, boundary.maxTwips);
  }
  if (marker === 'leftMargin') {
    const max = pageWidthTwips.value - rightMarginTwips.value - 720;
    return clampRounded(positionTwips, 0, max);
  }
  if (marker === 'rightMargin') {
    const fromRight = pageWidthTwips.value - positionTwips;
    const max = pageWidthTwips.value - leftMarginTwips.value - 720;
    return clampRounded(fromRight, 0, max);
  }
  if (marker === 'firstLineIndent') {
    const base = indentAreaLeftTwips.value + props.indentLeft;
    const indentFromBase = positionTwips - base;
    const max = Math.max(
      -props.indentLeft,
      indentAreaWidthTwips.value - props.indentLeft - props.indentRight - 720
    );
    return clampRounded(indentFromBase, -props.indentLeft, max);
  }
  if (marker === 'leftIndent') {
    const indentFromMargin = positionTwips - indentAreaLeftTwips.value;
    const max = Math.max(0, indentAreaWidthTwips.value - props.indentRight - 720);
    return clampRounded(indentFromMargin, 0, max);
  }
  // rightIndent
  const rightEdge = indentAreaRightTwips.value;
  const indentFromRight = rightEdge - positionTwips;
  const max = Math.max(0, indentAreaWidthTwips.value - props.indentLeft - 720);
  return clampRounded(indentFromRight, 0, max);
}

function emitChange(marker: MarkerType, value: number) {
  switch (marker) {
    case 'leftMargin':
      emit('left-margin-change', value);
      break;
    case 'rightMargin':
      emit('right-margin-change', value);
      break;
    case 'firstLineIndent':
      emit('first-line-indent-change', value);
      break;
    case 'leftIndent':
      emit('indent-left-change', value);
      break;
    case 'rightIndent':
      emit('indent-right-change', value);
      break;
    case 'tableBoundary':
      if (activeTableBoundary.value) {
        emit('table-boundary-change', activeTableBoundary.value, value);
      }
      break;
  }
}

onBeforeUnmount(() => {
  document.removeEventListener('mousemove', handleMove);
  document.removeEventListener('mouseup', handleUp);
});
</script>

<style scoped>
.docx-horizontal-ruler {
  display: block;
  position: relative;
  flex-shrink: 0;
}
.docx-horizontal-ruler__ticks {
  position: absolute;
  inset: 0;
  pointer-events: none;
}
/* Mirror React RulerTick: line at bottom, 1px wide, color via CSS var
   (--doc-text-subtle); falls back to #9aa0a6 if --doc-text-subtle is unset
   in the consumer scope. */
.docx-horizontal-ruler__tick-line {
  position: absolute;
  bottom: 0;
  width: 1px;
  background-color: var(--doc-text-subtle, #9aa0a6);
}
.docx-horizontal-ruler__tick-label {
  position: absolute;
  top: 3px;
  transform: translateX(-50%);
  font-size: 9px;
  color: var(--doc-text-muted, #5f6368);
  font-family: sans-serif;
  white-space: nowrap;
}
.docx-horizontal-ruler__tab {
  position: absolute;
  bottom: 0;
  width: 10px;
  height: 12px;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 8px;
  font-weight: 700;
  color: #555;
  cursor: pointer;
  user-select: none;
  transform: translateX(-5px);
}
.docx-ruler-table-boundary {
  user-select: none;
}
.docx-ruler-table-boundary__line {
  position: absolute;
  left: 50%;
  top: 2px;
  bottom: 2px;
  border-left: 1px dotted #b9c0c8;
}
.docx-horizontal-ruler__tooltip {
  position: absolute;
  top: -22px;
  transform: translateX(-50%);
  background: #333;
  color: #fff;
  font-size: 10px;
  font-family: sans-serif;
  padding: 2px 6px;
  border-radius: 3px;
  white-space: nowrap;
  pointer-events: none;
  z-index: 20;
}
</style>
