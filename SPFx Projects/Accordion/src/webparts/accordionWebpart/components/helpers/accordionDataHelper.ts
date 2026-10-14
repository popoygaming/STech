import { DEFAULT_ACCORDION_ITEMS_JSON } from '../constants/accordionDefaults';
import { IAccordionItem } from '../models/IAccordionItem';

export const getAccordionItems = (itemsJson: string | undefined): IAccordionItem[] => {
  try {
    const sourceJson: string = itemsJson && itemsJson.trim() ? itemsJson : DEFAULT_ACCORDION_ITEMS_JSON;
    const parsedData: unknown = JSON.parse(sourceJson);

    if (!Array.isArray(parsedData)) {
      return [];
    }

    const validItems: IAccordionItem[] = parsedData
      .filter((item: unknown): item is Record<string, unknown> => typeof item === 'object' && item !== null)
      .map((item: Record<string, unknown>) => ({
        title: String(item.title ?? '').trim(),
        content: String(item.content ?? '').trim()
      }));

    return validItems;
  } catch {
    return [];
  }
};
