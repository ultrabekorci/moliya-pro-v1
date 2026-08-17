/**
 * Amallar tarixi. Moliyaviy tizimda "kim, qachon, nimani o'zgartirdi" savoliga
 * javob bo'lishi shart — v1 da bunday jurnal umuman yo'q edi.
 */

import { auditTable } from '../db.js';
import { newId } from '../sheets.js';
import type { AuditEntry, Page } from '../../shared/types.js';
import type { StoredUser } from '../db.js';
import { normalizePaging } from '../../shared/validation.js';

const MAX_DETAILS_LENGTH = 900;

export function record(
  user: Pick<StoredUser, 'id' | 'login'>,
  action: string,
  entity: string,
  entityId: string,
  details: unknown = null,
): void {
  let text = '';
  if (typeof details === 'string') text = details;
  else if (details !== null && details !== undefined) {
    try {
      text = JSON.stringify(details);
    } catch {
      text = String(details);
    }
  }
  if (text.length > MAX_DETAILS_LENGTH) text = `${text.slice(0, MAX_DETAILS_LENGTH)}…`;

  auditTable.insert({
    id: newId(),
    at: new Date().toISOString(),
    userId: user.id,
    userLogin: user.login,
    action,
    entity,
    entityId,
    details: text,
  });
}

export function list(offset: unknown, limit: unknown): Page<AuditEntry> {
  const paging = normalizePaging(offset, limit);
  const all = auditTable.all().slice().sort((a, b) => (a.at < b.at ? 1 : a.at > b.at ? -1 : 0));
  return {
    items: all.slice(paging.offset, paging.offset + paging.limit),
    total: all.length,
    offset: paging.offset,
    limit: paging.limit,
  };
}
