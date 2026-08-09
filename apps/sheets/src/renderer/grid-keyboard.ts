import { CommandType, ICommandService, type IDisposable } from '@univerjs/core'
import { whenSheetEditorFocused } from '@univerjs/sheets-ui'
import { IShortcutService } from '@univerjs/ui'

import type { UniverRuntime } from './univer-state'

const PAGE_UP = 33
const PAGE_DOWN = 34
const END = 35
const HOME = 36
const PAGE_ROWS = 30

export const GridNavigationCommand = {
  Home: 'genoffice.sheet.operation.navigate-home',
  End: 'genoffice.sheet.operation.navigate-end',
  PageUp: 'genoffice.sheet.operation.navigate-page-up',
  PageDown: 'genoffice.sheet.operation.navigate-page-down',
} as const

type GridNavigationCommandId = (typeof GridNavigationCommand)[keyof typeof GridNavigationCommand]

function navigate(runtime: UniverRuntime, command: GridNavigationCommandId): boolean {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const selection = worksheet?.getSelection()
  const activeRange = selection?.getActiveRange()
  if (!workbook || !worksheet || !selection || !activeRange) return false

  const row = activeRange.getRow()
  const column = activeRange.getColumn()
  let target = null
  switch (command) {
    case GridNavigationCommand.Home:
      target = worksheet.getRange(row, 0)
      break
    case GridNavigationCommand.End:
      target =
        selection.getNextDataRange(runtime.univerAPI.Enum.Direction.RIGHT) ??
        worksheet.getRange(row, Math.max(0, worksheet.getMaxColumns() - 1))
      break
    case GridNavigationCommand.PageUp:
      target = worksheet.getRange(Math.max(0, row - PAGE_ROWS), column)
      break
    case GridNavigationCommand.PageDown:
      target = worksheet.getRange(
        Math.min(Math.max(0, worksheet.getMaxRows() - 1), row + PAGE_ROWS),
        column,
      )
      break
  }
  workbook.setActiveRange(target)
  worksheet.scrollToCell(target.getRow(), target.getColumn())
  return true
}

/**
 * Registers grid navigation with Univer's own shortcut dispatcher. This shares
 * its capture listener and its editor context predicates, avoiding races with
 * the built-in Delete/Backspace clear shortcut.
 */
export function installGridKeyboardShortcuts(runtime: UniverRuntime): IDisposable {
  const injector = runtime.univer.__getInjector()
  const commandService = injector.get(ICommandService)
  const shortcutService = injector.get(IShortcutService)
  const shortcuts: readonly { id: GridNavigationCommandId; binding: number }[] = [
    { id: GridNavigationCommand.Home, binding: HOME },
    { id: GridNavigationCommand.End, binding: END },
    { id: GridNavigationCommand.PageUp, binding: PAGE_UP },
    { id: GridNavigationCommand.PageDown, binding: PAGE_DOWN },
  ]
  const disposables = [
    ...shortcuts.map(({ id }) =>
      commandService.registerCommand({
        id,
        type: CommandType.OPERATION,
        handler: () => navigate(runtime, id),
      }),
    ),
    ...shortcuts.map(({ id, binding }) =>
      shortcutService.registerShortcut({
        id,
        binding,
        preconditions: whenSheetEditorFocused,
        priority: 1,
      }),
    ),
  ]
  return {
    dispose: () => disposables.forEach((disposable) => disposable.dispose()),
  }
}
