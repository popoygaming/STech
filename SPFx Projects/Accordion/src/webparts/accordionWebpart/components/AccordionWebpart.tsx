import * as React from 'react';
import * as DOMPurifyModule from 'dompurify';
import { DefaultButton, Icon, Text } from '@fluentui/react';
import { RichText } from '@pnp/spfx-controls-react/lib/RichText';
import type { IAccordionWebpartProps } from './IAccordionWebpartProps';
import type { IAccordionItem } from './models/IAccordionItem';
import styles from './AccordionWebpart.module.scss';

interface IAccordionWebpartState {
  openIndices: number[];
  draftItems: IAccordionItem[];
}

export default class AccordionWebpart extends React.Component<IAccordionWebpartProps, IAccordionWebpartState> {
  public constructor(props: IAccordionWebpartProps) {
    super(props);

    this.state = {
      openIndices: props.isEditMode ? props.items.map((_, index) => index) : [],
      draftItems: props.items.map((item: IAccordionItem) => ({ ...item }))
    };
  }

  private onToggleContainer = (index: number): void => {
    this.setState((previousState: IAccordionWebpartState) => ({
      openIndices: previousState.openIndices.indexOf(index) > -1
        ? previousState.openIndices.filter((itemIndex: number) => itemIndex !== index)
        : [...previousState.openIndices, index]
    }));
  };

  private onRichTextChange = (index: number, newValue: string): string => {
    this.setState((previousState: IAccordionWebpartState) => {
      const draftItems: IAccordionItem[] = previousState.draftItems.map((item: IAccordionItem, itemIndex: number) => (
        itemIndex === index ? { ...item, content: newValue } : item
      ));

      this.props.onInlineSave(draftItems);

      return { draftItems };
    });

    return newValue;
  };

  private onHeaderTextChange = (index: number, event: React.ChangeEvent<HTMLInputElement>): void => {
    const newHeaderText = event.target.value;

    this.setState((previousState: IAccordionWebpartState) => {
      const draftItems: IAccordionItem[] = previousState.draftItems.map((item: IAccordionItem, itemIndex: number) => (
        itemIndex === index ? { ...item, title: newHeaderText } : item
      ));

      this.props.onInlineSave(draftItems);

      return { draftItems };
    });
  };

  private onAddAccordion = (): void => {
    this.setState((previousState: IAccordionWebpartState) => {
      const draftItems: IAccordionItem[] = [
        ...previousState.draftItems,
        {
          title: 'New section',
          content: '<p style="color:#000000;">Add content here.</p>'
        }
      ];

      this.props.onInlineSave(draftItems);

      return {
        draftItems,
        openIndices: [...previousState.openIndices, draftItems.length - 1]
      };
    });
  };

  private onDeleteAccordion = (index: number): void => {
    this.setState((previousState: IAccordionWebpartState) => {
      const draftItems: IAccordionItem[] = previousState.draftItems.filter((_, itemIndex: number) => itemIndex !== index);

      const openIndices: number[] = previousState.openIndices
        .filter((itemIndex: number) => itemIndex !== index)
        .map((itemIndex: number) => (itemIndex > index ? itemIndex - 1 : itemIndex));

      this.props.onInlineSave(draftItems);

      return {
        draftItems,
        openIndices
      };
    });
  };

  public componentDidUpdate(prevProps: IAccordionWebpartProps): void {
    if (prevProps.items !== this.props.items) {
      this.setState({ draftItems: this.props.items.map((item: IAccordionItem) => ({ ...item })) });
    }

    if (!prevProps.isEditMode && this.props.isEditMode) {
      this.setState({
        openIndices: this.props.items.map((_, index) => index)
      });
    }
  }

  private sanitizeRichText = (html: string): string => {
    const options = {
      USE_PROFILES: { html: true },
      FORBID_TAGS: ['script', 'style', 'iframe', 'object', 'embed'],
      ADD_ATTR: ['target', 'rel']
    };

    const purifierModule = DOMPurifyModule as unknown as {
      sanitize?: (dirty: string, config?: unknown) => string;
      default?:
        | { sanitize?: (dirty: string, config?: unknown) => string }
        | ((window: Window) => { sanitize: (dirty: string, config?: unknown) => string });
    };

    if (typeof purifierModule.sanitize === 'function') {
      return purifierModule.sanitize(html, options);
    }

    if (typeof purifierModule.default === 'object' && typeof purifierModule.default.sanitize === 'function') {
      return purifierModule.default.sanitize(html, options);
    }

    if (typeof window !== 'undefined' && typeof purifierModule.default === 'function') {
      const purifier = purifierModule.default(window);
      return purifier.sanitize(html, options);
    }

    return html;
  };

  public render(): React.ReactElement<IAccordionWebpartProps> {
    const { items, textSize, isDarkTheme, isEditMode } = this.props;
    const { openIndices, draftItems } = this.state;
    const themeClass = isDarkTheme ? styles.dark : styles.light;
    const sizeClass = textSize === 'small' ? styles.textSmall : textSize === 'large' ? styles.textLarge : styles.textMedium;
    const visibleItems: IAccordionItem[] = isEditMode
      ? draftItems
      : items.filter((item: IAccordionItem) => item.title.trim().length > 0 && item.content.trim().length > 0);

    return (
      <section className={`${styles.accordionWebpart} ${themeClass} ${sizeClass}`}>
        {isEditMode && (
          <div className={styles.editorActionsBar}>
            <DefaultButton text="Add accordion" onClick={this.onAddAccordion} />
          </div>
        )}

        {visibleItems.map((item: IAccordionItem, index: number) => {
          const isOpen = openIndices.indexOf(index) > -1;
          const contentId = `accordion-content-${index}`;
          const safeContent = this.sanitizeRichText(item.content);

          return (
            <article key={`accordion-${index}`} className={`${styles.itemCard} ${isOpen ? styles.itemCardActive : ''}`}>
              <button
                type="button"
                className={styles.itemButton}
                onClick={() => this.onToggleContainer(index)}
                aria-expanded={isOpen}
                aria-controls={contentId}
              >
                {isEditMode ? (
                  <input
                    type="text"
                    value={item.title}
                    className={styles.itemTitleInput}
                    onClick={(event) => event.stopPropagation()}
                    onFocus={(event) => event.stopPropagation()}
                    onChange={(event) => this.onHeaderTextChange(index, event)}
                    placeholder="Section Header"
                    aria-label={`Accordion header ${index + 1}`}
                  />
                ) : (
                  <Text variant="mediumPlus" className={styles.itemTitle}>{item.title}</Text>
                )}
                <Icon iconName={isOpen ? 'ChevronUp' : 'ChevronDown'} className={styles.indicator} />
              </button>

              {isOpen && (
                <div
                  id={contentId}
                  className={styles.itemContentWrap}
                >
                  {isEditMode ? (
                    <div className={styles.editorContentWrap}>
                      <div className={styles.itemEditActions}>
                        <DefaultButton
                          text="Delete accordion"
                          onClick={() => this.onDeleteAccordion(index)}
                        />
                      </div>
                      <RichText
                        value={item.content}
                        isEditMode={true}
                        onChange={(newValue) => this.onRichTextChange(index, newValue)}
                      />
                    </div>
                  ) : (
                    <div
                      className={styles.itemContent}
                      dangerouslySetInnerHTML={{ __html: safeContent }}
                    />
                  )}
                </div>
              )}
            </article>
          );
        })}
      </section>
    );
  }
}
