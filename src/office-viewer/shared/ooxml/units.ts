export const EMU_PER_INCH = 914400;
export const PX_PER_INCH = 96;

export function emuToPx(emu: number) {
  return (emu / EMU_PER_INCH) * PX_PER_INCH;
}
