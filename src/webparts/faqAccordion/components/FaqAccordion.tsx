import * as React from 'react';
import styles from './FaqAccordion.module.scss';
import { IFaqAccordionProps } from './IFaqAccordionProps';
import { IFaqItem, IFaqCategoryGroup } from '../models/IFaqItem';
import { SPFaqService } from '../services/SPFaqService';

import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';

// Plain-object "set" types. Using these instead of ES2015 Set/Map keeps the
// component compatible with the default SPFx TypeScript lib settings.
interface IStringSet { [key: string]: boolean; }
interface INumberSet { [key: number]: boolean; }

const FaqAccordion: React.FC<IFaqAccordionProps> = (props) => {
  const { anchorId, siteUrl, listId, setId, context, title, displayMode, updateTitle, styles: styleOverrides } = props;

  // Sanitize the anchor ID — strip leading '#' if the user copy-pasted one,
  // and trim whitespace. We don't enforce stricter rules (letters/digits/dashes)
  // here because browsers actually accept most characters in id attributes.
  const safeAnchorId = (anchorId || '').trim().replace(/^#+/, '');

  // Build inline CSS variables from the property-pane values. Empty strings
  // fall through to the SCSS fallbacks so the web part still looks right
  // before the user has configured anything.
  const cssVars: React.CSSProperties = {};
  const cssVarMap: { [k: string]: string | undefined } = {
    '--faq-category-bg': styleOverrides && styleOverrides.categoryBgColor,
    '--faq-category-color': styleOverrides && styleOverrides.categoryTextColor,
    '--faq-category-font': styleOverrides && styleOverrides.categoryFont,
    '--faq-category-weight': styleOverrides && styleOverrides.categoryFontWeight,
    '--faq-question-bg': styleOverrides && styleOverrides.questionBgColor,
    '--faq-question-color': styleOverrides && styleOverrides.questionTextColor,
    '--faq-question-font': styleOverrides && styleOverrides.questionFont,
    '--faq-question-weight': styleOverrides && styleOverrides.questionFontWeight
  };
  for (const key in cssVarMap) {
    if (Object.prototype.hasOwnProperty.call(cssVarMap, key)) {
      const val = cssVarMap[key];
      if (val) {
        // React's typing for CSSProperties doesn't include custom properties,
        // hence the cast. At runtime it sets the inline custom property correctly.
        (cssVars as { [k: string]: string })[key] = val;
      }
    }
  }

  const [items, setItems] = React.useState<IFaqItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [search, setSearch] = React.useState<string>('');
  const [openCategories, setOpenCategories] = React.useState<IStringSet>({});
  const [openQuestions, setOpenQuestions] = React.useState<INumberSet>({});

  // Fetch items whenever the selected list changes
  React.useEffect(() => {
    let cancelled = false;

    const load = async (): Promise<void> => {
      if (!listId) {
        setItems([]);
        return;
      }
      setLoading(true);
      setError(undefined);
      try {
        const service = new SPFaqService(context);
        const data = await service.getFaqItems(listId, siteUrl, setId);
        if (!cancelled) {
          setItems(data);
          // reset expansion state when data reloads
          setOpenCategories({});
          setOpenQuestions({});
        }
      } catch (ex) {
        if (!cancelled) {
          setError(ex instanceof Error ? ex.message : String(ex));
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    };

    load().catch(() => { /* errors handled inside load() via setError */ });
    return () => {
      cancelled = true;
    };
  }, [listId, siteUrl, setId, context]);

  // Group + sort alphabetically by category and then by question.
  // Uses plain objects and arrays only — no Map/Set — to keep the ES5 lib happy.
  const groups: IFaqCategoryGroup[] = React.useMemo(() => {
    const needle = search.trim().toLowerCase();

    const filtered: IFaqItem[] = needle
      ? items.filter((i: IFaqItem) =>
          i.question.toLowerCase().indexOf(needle) >= 0 ||
          i.category.toLowerCase().indexOf(needle) >= 0 ||
          stripHtml(i.answer).toLowerCase().indexOf(needle) >= 0)
      : items;

    const byCategory: { [key: string]: IFaqItem[] } = {};
    const categoryNames: string[] = [];

    for (let i = 0; i < filtered.length; i++) {
      const item = filtered[i];
      if (!byCategory[item.category]) {
        byCategory[item.category] = [];
        categoryNames.push(item.category);
      }
      byCategory[item.category].push(item);
    }

    categoryNames.sort((a: string, b: string) =>
      a.localeCompare(b, undefined, { sensitivity: 'base' }));

    const sorted: IFaqCategoryGroup[] = [];
    for (let j = 0; j < categoryNames.length; j++) {
      const cat = categoryNames[j];
      // Sort by SortOrder ascending. Items without a SortOrder go to the end.
      // Ties (same SortOrder, or both missing) fall back to alphabetical.
      const list = byCategory[cat].slice().sort((a: IFaqItem, b: IFaqItem) => {
        const aHas = a.sortOrder !== undefined;
        const bHas = b.sortOrder !== undefined;
        if (aHas && bHas) {
          const diff = (a.sortOrder as number) - (b.sortOrder as number);
          if (diff !== 0) return diff;
        } else if (aHas) {
          return -1;
        } else if (bHas) {
          return 1;
        }
        return a.question.localeCompare(b.question, undefined, { sensitivity: 'base' });
      });
      sorted.push({ category: cat, items: list });
    }
    return sorted;
  }, [items, search]);

  // Auto-expand all categories when the user is searching so matches are visible
  const isSearching = search.trim().length > 0;
  const effectiveOpenCategories: IStringSet = React.useMemo(() => {
    if (!isSearching) return openCategories;
    const all: IStringSet = {};
    for (let i = 0; i < groups.length; i++) {
      all[groups[i].category] = true;
    }
    return all;
  }, [isSearching, openCategories, groups]);

  const toggleCategory = (name: string): void => {
    setOpenCategories((prev: IStringSet) => {
      const next: IStringSet = {};
      // copy existing keys
      for (const k in prev) {
        if (Object.prototype.hasOwnProperty.call(prev, k)) {
          next[k] = prev[k];
        }
      }
      if (next[name]) {
        delete next[name];
      } else {
        next[name] = true;
      }
      return next;
    });
  };

  const toggleQuestion = (id: number): void => {
    setOpenQuestions((prev: INumberSet) => {
      const next: INumberSet = {};
      for (const k in prev) {
        if (Object.prototype.hasOwnProperty.call(prev, k)) {
          next[k as unknown as number] = prev[k as unknown as number];
        }
      }
      if (next[id]) {
        delete next[id];
      } else {
        next[id] = true;
      }
      return next;
    });
  };

  const onKeyToggle = (e: React.KeyboardEvent, toggleFn: () => void): void => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      toggleFn();
    }
  };

  return (
    <div
      id={safeAnchorId || undefined}
      className={styles.faqAccordion}
      style={cssVars}
    >
      <WebPartTitle
        displayMode={displayMode}
        title={title}
        updateProperty={updateTitle}
      />

      {!listId && (
        <MessageBar messageBarType={MessageBarType.info}>
          Open the web part properties and choose the SharePoint list that contains your FAQ items.
        </MessageBar>
      )}

      {listId && (
        <div className={styles.searchBox}>
          <SearchBox
            placeholder="Search FAQs..."
            value={search}
            onChange={(_, newValue) => setSearch(newValue || '')}
            onClear={() => setSearch('')}
          />
        </div>
      )}

      {loading && (
        <div className={styles.loading}>
          <Spinner size={SpinnerSize.medium} label="Loading FAQs..." />
        </div>
      )}

      {error && (() => {
        const friendly = friendlyError(error);
        return (
          <MessageBar messageBarType={MessageBarType.error} isMultiline>
            <div style={{ fontWeight: 600, marginBottom: 4 }}>
              {friendly ? friendly.title : 'Could not load FAQ items'}
            </div>
            {friendly && (
              <div style={{ marginBottom: 6 }}>{friendly.message}</div>
            )}
            {friendly && friendly.missingField && (
              <div style={{ marginBottom: 6 }}>
                Required columns for this web part:
                <ul style={{ margin: '4px 0 4px 18px', padding: 0 }}>
                  <li><strong>Question</strong> — Single line of text</li>
                  <li><strong>Answer</strong> — Multiple lines of text (Enhanced rich text)</li>
                  <li><strong>Category</strong> — Choice</li>
                  <li><strong>SortOrder</strong> — Number (optional)</li>
                </ul>
              </div>
            )}
            <details>
              <summary style={{ cursor: 'pointer' }}>Technical details</summary>
              <code style={{ display: 'block', whiteSpace: 'pre-wrap', marginTop: 6, fontSize: 12 }}>
                {error}
              </code>
            </details>
          </MessageBar>
        );
      })()}

      {!loading && !error && listId && groups.length === 0 && (
        <div className={styles.empty}>
          {isSearching ? 'No FAQs match your search.' : 'No FAQ items found in this list.'}
        </div>
      )}

      {!loading && !error && groups.map((group: IFaqCategoryGroup) => {
        const isOpen = !!effectiveOpenCategories[group.category];
        const categoryId = `cat-${slug(group.category)}`;
        return (
          <div key={group.category} className={styles.category}>
            <button
              type="button"
              className={styles.categoryHeader}
              aria-expanded={isOpen}
              aria-controls={categoryId}
              onClick={() => toggleCategory(group.category)}
              onKeyDown={(e) => onKeyToggle(e, () => toggleCategory(group.category))}
            >
              <Icon
                iconName="ChevronRight"
                className={`${styles.chevron} ${isOpen ? styles.chevronOpen : ''}`}
              />
              <span>{group.category}</span>
              <span className={styles.countBadge}>
                {group.items.length} {group.items.length === 1 ? 'question' : 'questions'}
              </span>
            </button>

            {isOpen && (
              <ul id={categoryId} className={styles.questionsList}>
                {group.items.map((item: IFaqItem) => {
                  const qOpen = !!openQuestions[item.id];
                  const questionId = `q-${item.id}`;
                  return (
                    <li key={item.id} className={styles.question}>
                      <button
                        type="button"
                        className={styles.questionHeader}
                        aria-expanded={qOpen}
                        aria-controls={questionId}
                        onClick={() => toggleQuestion(item.id)}
                        onKeyDown={(e) => onKeyToggle(e, () => toggleQuestion(item.id))}
                      >
                        <Icon
                          iconName="ChevronRight"
                          className={`${styles.chevron} ${qOpen ? styles.chevronOpen : ''}`}
                        />
                        <span>{item.question}</span>
                      </button>
                      {qOpen && (
                        <div
                          id={questionId}
                          className={styles.answer}
                          // Answer is rich text HTML stored by SharePoint. It is sanitized
                          // server-side when saved via the SharePoint UI.
                          dangerouslySetInnerHTML={{ __html: item.answer }}
                        />
                      )}
                    </li>
                  );
                })}
              </ul>
            )}
          </div>
        );
      })}
    </div>
  );
};

function stripHtml(html: string): string {
  if (!html) return '';
  return html.replace(/<[^>]*>/g, ' ');
}

interface IFriendlyError {
  title: string;
  message: string;
  missingField?: string;
}

/**
 * Translates raw SharePoint error strings into human-readable guidance.
 * Returns undefined if the error doesn't match a known pattern.
 */
function friendlyError(raw: string): IFriendlyError | undefined {
  if (!raw) return undefined;

  // Missing column on the list.
  const missingFieldMatch = raw.match(/The field or property '([^']+)' does not exist/i);
  if (missingFieldMatch) {
    const field = missingFieldMatch[1];
    return {
      title: "This list isn't set up for the FAQ web part",
      message: `The list is missing a "${field}" column. Either choose a different list or add the required columns to this one.`,
      missingField: field
    };
  }

  // List GUID not found (deleted, renamed, or you don't have access).
  if (/List does not exist/i.test(raw) || /\(404\)/.test(raw)) {
    return {
      title: 'List not found',
      message: 'The selected list could not be opened. It may have been deleted, moved, or you may not have permission to read it. Pick a different list in the web part properties.'
    };
  }

  // Cross-site access denied.
  if (/\(401\)/.test(raw) || /\(403\)/.test(raw) || /unauthorized/i.test(raw)) {
    return {
      title: 'Access denied',
      message: 'You do not have permission to read the selected list. Ask the site owner to grant you Read access, or pick a different list.'
    };
  }

  // Bad site URL pasted into the property pane.
  if (/\bbadrequest\b/i.test(raw) && /\(400\)/.test(raw)) {
    return {
      title: 'The request was rejected by SharePoint',
      message: 'This usually means the site URL or list selection is invalid. Double-check the Site URL property and re-pick the list.'
    };
  }

  return undefined;
}

function slug(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
}

export default FaqAccordion;
