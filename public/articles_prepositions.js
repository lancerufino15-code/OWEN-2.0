/**
 * Client-side stoplist and normalization helpers for articles/prepositions.
 *
 * Used by: `public/chat.js` to filter highlight tokens and metadata.
 *
 * Key exports:
 * - `PREPOSITIONS_SINGLE_BASE`, `PREPOSITIONS_MULTI_BASE`: canonical token lists.
 * - `ARTICLES_AND_PREPOSITIONS_SET`, `MULTIWORD_PREPOSITION_REGEX`: lookup helpers.
 * - `isArticleOrPreposition`: normalization + membership check.
 *
 * Assumptions:
 * - Normalization is case-insensitive, strips diacritics, and flattens dashes/apostrophes to spaces.
 */

const ARTICLES_BASE = ["a", "an", "the"];

/** Single-word English prepositions (plus a few project-specific stopwords). */
export const PREPOSITIONS_SINGLE_BASE = [
  "aboard",
  "about",
  "above",
  "across",
  "after",
  "against",
  "along",
  "alongside",
  "amid",
  "amidst",
  "among",
  "amongst",
  "around",
  "as",
  "astride",
  "at",
  "atop",
  "bar",
  "before",
  "behind",
  "below",
  "beneath",
  "beside",
  "besides",
  "between",
  "betwixt",
  "beyond",
  "but",
  "by",
  "circa",
  "concerning",
  "considering",
  "despite",
  "down",
  "during",
  "except",
  "excluding",
  "failing",
  "following",
  "for",
  "from",
  "given",
  "in",
  "inside",
  "into",
  "less",
  "like",
  "minus",
  "near",
  "notwithstanding",
  "of",
  "off",
  "on",
  "onto",
  "opposite",
  "out",
  "outside",
  "over",
  "past",
  "pending",
  "per",
  "plus",
  "re",
  "regarding",
  "respecting",
  "round",
  "save",
  "sans",
  "since",
  "than",
  "through",
  "throughout",
  "till",
  "to",
  "toward",
  "towards",
  "under",
  "underneath",
  "unlike",
  "until",
  "up",
  "upon",
  "versus",
  "via",
  "with",
  "within",
  "without",
  "worth",
];

/** Multi-word / compound prepositions and prepositional phrases. */
export const PREPOSITIONS_MULTI_BASE = [
  "according to",
  "ahead of",
  "along with",
  "apart from",
  "as for",
  "as of",
  "as per",
  "as regards",
  "as to",
  "aside from",
  "back of",
  "because of",
  "by dint of",
  "by means of",
  "by reason of",
  "by virtue of",
  "by way of",
  "close to",
  "contrary to",
  "depending on",
  "due to",
  "except for",
  "far from",
  "for lack of",
  "for the sake of",
  "in accordance with",
  "in addition to",
  "in advance of",
  "in back of",
  "in case of",
  "in comparison with",
  "in compliance with",
  "in connection with",
  "in contrast to",
  "in exchange for",
  "in favor of",
  "in favour of",
  "in front of",
  "in keeping with",
  "in light of",
  "in line with",
  "in lieu of",
  "in place of",
  "in regard to",
  "in regards to",
  "in relation to",
  "in respect of",
  "in response to",
  "in search of",
  "in spite of",
  "in support of",
  "in terms of",
  "in the event of",
  "in the face of",
  "in the middle of",
  "in the name of",
  "in the vicinity of",
  "in the wake of",
  "in view of",
  "inside of",
  "instead of",
  "near to",
  "next to",
  "on account of",
  "on behalf of",
  "on the basis of",
  "on the part of",
  "on the strength of",
  "on the verge of",
  "on top of",
  "out of",
  "outside of",
  "owing to",
  "prior to",
  "pursuant to",
  "regardless of",
  "relative to",
  "subsequent to",
  "thanks to",
  "together with",
  "up to",
  "with a view to",
  "with reference to",
  "with regard to",
  "with regards to",
  "with respect to",
  "with the aid of",
  "with the exception of",
  "with the help of",
  "vis a vis",
];

const normalize = (value) =>
  value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/['\u2019]/g, " ")
    .replace(/[-\u2013\u2014]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const buildStopSet = () => {
  const set = new Set();
  const addAll = (list) => {
    list.forEach((item) => {
      const normalized = normalize(item);
      if (normalized) set.add(normalized);
    });
  };
  addAll(ARTICLES_BASE);
  addAll(PREPOSITIONS_SINGLE_BASE);
  addAll(PREPOSITIONS_MULTI_BASE);
  return set;
};

const STOP_SET = buildStopSet();

/**
 * Normalized stopword set for articles and prepositions.
 */
export const ARTICLES_AND_PREPOSITIONS_SET = STOP_SET;

const escapeRegex = (value) => value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/\s+/g, "\\s+");

/**
 * Regex that matches normalized multi-word prepositions.
 */
export const MULTIWORD_PREPOSITION_REGEX = new RegExp(
  `\\b(?:${PREPOSITIONS_MULTI_BASE.map(normalize).map(escapeRegex).join("|")})\\b`,
  "gi",
);

/**
 * Check whether a token or phrase is an article/preposition after normalization.
 *
 * @param tokenOrPhrase - Raw token or phrase.
 * @returns True when the normalized form is in the stopword set.
 */
export function isArticleOrPreposition(tokenOrPhrase) {
  const normalized = normalize(tokenOrPhrase || "");
  if (!normalized) return false;
  return STOP_SET.has(normalized);
}
