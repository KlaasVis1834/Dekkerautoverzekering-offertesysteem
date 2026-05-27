-- Repair existing offer delivery statuses without changing business data.
-- Run this manually in Supabase SQL Editor after the app has deployed.
--
-- The Flask app itself updates delivery_status during write actions. This file
-- only repairs old rows where files or aanvraag fields already prove that an
-- offer was processed.

ALTER TABLE public.offers
ADD COLUMN IF NOT EXISTS graph_message_id TEXT;

UPDATE public.offers
SET delivery_status = CASE
    WHEN COALESCE(aanvraag_status, '') = 'afgehandeld'
        THEN 'afgehandeld'
    WHEN COALESCE(aanvraag_ontvangen_at, '') <> ''
        THEN 'aanvraag_ontvangen'
    WHEN COALESCE(graph_message_id, '') <> ''
        THEN 'outlook_concept_klaar'
    WHEN COALESCE(eml_path, '') <> ''
        THEN 'email_klaar'
    WHEN COALESCE(post_letter_path, '') <> ''
        THEN 'post_klaar'
    WHEN COALESCE(offer_pdf_path, '') <> '' AND COALESCE(email, '') <> ''
        THEN 'email_klaar'
    WHEN COALESCE(offer_pdf_path, '') <> ''
        THEN 'post_klaar'
    ELSE COALESCE(delivery_status, 'nieuw')
END,
updated_at = COALESCE(updated_at, to_char(now(), 'YYYY-MM-DD HH24:MI:SS'))
WHERE COALESCE(delivery_status, 'nieuw') IN ('nieuw', 'verstuurd')
  AND (
      COALESCE(aanvraag_status, '') = 'afgehandeld'
      OR COALESCE(aanvraag_ontvangen_at, '') <> ''
      OR COALESCE(graph_message_id, '') <> ''
      OR COALESCE(eml_path, '') <> ''
      OR COALESCE(post_letter_path, '') <> ''
      OR COALESCE(offer_pdf_path, '') <> ''
  );
