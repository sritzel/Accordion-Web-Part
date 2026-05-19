import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IFaqStyleOverrides {
  categoryBgColor: string;
  categoryTextColor: string;
  categoryFont: string;
  categoryFontWeight: string;
  questionBgColor: string;
  questionTextColor: string;
  questionFont: string;
  questionFontWeight: string;
}

export interface IFaqAccordionProps {
  title: string;
  anchorId: string;      // optional HTML id for in-page jump links
  siteUrl: string;       // absolute URL of the site containing the list; blank = current site
  listId: string;
  setId: string;         // ID of the Set to filter by; '' = no filter
  context: WebPartContext;
  isDarkTheme: boolean;
  displayMode: DisplayMode;
  updateTitle: (value: string) => void;
  styles: IFaqStyleOverrides;
}
