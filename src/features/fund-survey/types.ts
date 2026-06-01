export type FundRow = {
  company: string;
  fund: string;
};

export type StatusTone = "success" | "info" | "";
export type Status = { message: string; tone: StatusTone };
