import type { SelectedRange } from "../excel/selection";

/** Turn a common incomplete cross-range plan failure into an actionable prompt. */
export function userFacingError(error: unknown, selected?: Pick<SelectedRange, "sheet" | "address">): string {
  const message = error instanceof Error ? error.message : "操作失败";
  if (!/写入数据与目标区域的行列数不一致|数据每行列数必须一致/.test(message)) return message;
  const target = selected ? `${selected.sheet}!${selected.address}` : "当前选区";
  return `模型生成的数据与目标区域 ${target} 的尺寸不一致，本轮已停止且未修改工作簿。通常是目标区域与参考区域没有同时提供，或发送前在不同工作表依次选择后只保留了最后一个选区。请重新发送时一次性选择尽量完整的目标区域；需要参考另一张表时，在提示中写明参考工作表和连续地址，并在发送前确认两边范围完整。发送后切换工作表不会改变已捕获的选区。`;
}
