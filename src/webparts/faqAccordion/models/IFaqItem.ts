export interface IFaqItem {
  id: number;
  question: string;
  answer: string;   // rich text HTML
  category: string;
  sortOrder: number | undefined;
}

export interface IFaqCategoryGroup {
  category: string;
  items: IFaqItem[];
}
