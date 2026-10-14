import { PersonaPresence } from '@fluentui/react';
import type { IEmployeeProfile } from '../services/EmployeeDirectoryService';

export function getValidPageSize(pageSize?: number): number {
  if (!pageSize || pageSize < 1) {
    return 20;
  }

  return pageSize;
}

export function getDefaultSearchTextFromUrl(): string {
  try {
    const params: URLSearchParams = new URLSearchParams(window.location.search);
    const fromSearch: string | null = params.get('search');
    const fromSearchText: string | null = params.get('searchText');
    const fromQuery: string | null = params.get('q');

    return (fromSearch || fromSearchText || fromQuery || '').trim();
  } catch (_error) {
    return '';
  }
}

export function getFilteredEmployees(employees: IEmployeeProfile[], searchText: string): IEmployeeProfile[] {
  const query: string = searchText.trim().toLowerCase();

  if (!query) {
    return employees;
  }

  return employees.filter((employee: IEmployeeProfile) => {
    const fullName: string = (employee.fullName || '').toLowerCase();
    const email: string = (employee.email || '').toLowerCase();
    const role: string = (employee.role || '').toLowerCase();

    return fullName.indexOf(query) > -1 || email.indexOf(query) > -1 || role.indexOf(query) > -1;
  });
}

export function mapAvailabilityToPresence(availability?: string): PersonaPresence {
  switch ((availability || '').toLowerCase()) {
    case 'available':
      return PersonaPresence.online;
    case 'away':
    case 'be right back':
      return PersonaPresence.away;
    case 'busy':
    case 'in a call':
    case 'in a meeting':
    case 'do not disturb':
    case 'presenting':
      return PersonaPresence.dnd;
    case 'offline':
    case 'presence unknown':
      return PersonaPresence.offline;
    default:
      return PersonaPresence.none;
  }
}

export function toTitleCase(value?: string): string {
  if (!value) {
    return 'N/A';
  }

  return value
    .trim()
    .toLowerCase()
    .split(/\s+/)
    .map((word: string) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');
}

export function formatDateOnly(value?: string): string {
  if (!value) {
    return 'N/A';
  }

  const parsedDate: Date = new Date(value);

  if (isNaN(parsedDate.getTime())) {
    return value;
  }

  return parsedDate.toLocaleDateString();
}