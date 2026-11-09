export const STATUS_TONE: Record<string, "gray" | "yellow" | "navy" | "blue" | "red"> = {
  SUBMITTED: "yellow", MANAGER_APPROVED: "navy", HR_VERIFIED: "navy", APPROVED: "blue", REJECTED: "red", WITHDRAWN: "gray",
};
export const STATUS_TEXT: Record<string, string> = {
  SUBMITTED: "Awaiting manager review", MANAGER_APPROVED: "Awaiting HR verification", HR_VERIFIED: "Awaiting final approval",
  APPROVED: "Approved", REJECTED: "Rejected", WITHDRAWN: "Withdrawn",
};
