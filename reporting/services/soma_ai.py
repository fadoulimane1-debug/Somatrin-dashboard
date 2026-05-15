"""
SOMA AI — assistant conversationnel avec accent sur la confidentialité.

Principes :
- Pas d’appel à des API cloud externes par défaut : Ollama doit tourner en local (boucleman).
- Les messages utilisateur ne sont pas journalisés en clair (longueur / métadonnées uniquement).
- L’historique de conversation est stocké côté serveur (session), pas renvoyé au client pour réinjection.
- Aucune donnée de credential Odoo injectée dans les prompts.
"""
from __future__ import annotations

import ipaddress
import json
import logging
import re
import threading
import time
import urllib.error
import urllib.request
from typing import Any
from urllib.parse import urlparse

from django.conf import settings

logger = logging.getLogger(__name__)

_RATE_LOCK = threading.Lock()
_RATE_BUCKETS: dict[int, list[float]] = {}

SYSTEM_PROMPT_FR = """Tu es SOMA AI, assistant interne pour les collaborateurs SOMATRIN.

Règles strictes :
- Toute réponse affichée à l’utilisateur doit être entièrement en français (aucune phrase ou formule en anglais).
- Réponds de façon concise et professionnelle.
- Tu n’as pas accès direct aux bases Odoo ou ERP : n’invente jamais de chiffres, montants ni listes métier précises, sauf lorsqu’un second message système te fournit explicitement des indicateurs tableau de bord — dans ce cas tu cites uniquement ces valeurs pour ce qui les concerne.
- Hors de ces indicateurs fournis, pour les données opérationnelles oriente l’utilisateur vers l’écran SOMATRIN concerné (gasoil, achats, production, etc.).
- Ne demande jamais de mot de passe, jeton API ou donnée d’authentification ; ne répète pas ce type d’information si l’utilisateur en envoie.
- Si une question dépasse ton périmètre général (conseil juridique, médical, financier engageant), décline poliment et propose de passer par les équipes concernées.
"""


def _rate_limit_ok(user_id: int) -> bool:
    lim = getattr(settings, 'SOMA_AI_RATE_LIMIT_PER_MINUTE', 15)
    now = time.time()
    with _RATE_LOCK:
        bucket = _RATE_BUCKETS.setdefault(user_id, [])
        bucket[:] = [t for t in bucket if now - t < 60.0]
        if len(bucket) >= lim:
            return False
        bucket.append(now)
        return True


def _strip_control_chars(text: str) -> str:
    if '\x00' in text:
        text = text.replace('\x00', '')
    return ''.join(ch for ch in text if ch >= ' ' or ch in '\n\r\t')


def sanitize_user_message(raw: str, max_chars: int) -> tuple[str | None, str | None]:
    """
    Valide et nettoie le message utilisateur.
    Retourne (texte_ok, erreur_code) ; erreur_code pour messages courts côté API.
    """
    if not isinstance(raw, str):
        return None, 'message_invalide'
    text = raw.strip()
    text = _strip_control_chars(text)
    if not text:
        return None, 'message_vide'
    if len(text) > max_chars:
        return None, 'message_trop_long'
    return text, None


def _hostname_allowed(host: str) -> bool:
    host = (host or '').strip().lower()
    if host in ('127.0.0.1', 'localhost', '::1'):
        return True
    if getattr(settings, 'SOMA_AI_ALLOW_PRIVATE_LAN', False):
        try:
            ip = ipaddress.ip_address(host)
            return ip.is_private or ip.is_loopback
        except ValueError:
            return False
    return False


def validate_ollama_base_url(url: str) -> tuple[str | None, str | None]:
    """
    Évite les SSRF : URL Ollama doit pointer vers loopback (ou LAN privé si activé explicitement).
    Retourne (url_normalisée, erreur).
    """
    if not url or not isinstance(url, str):
        return None, 'ollama_url_manquante'
    url = url.strip()
    if not url.startswith(('http://', 'https://')):
        url = 'http://' + url
    parsed = urlparse(url)
    if parsed.scheme not in ('http', 'https'):
        return None, 'ollama_scheme_invalide'
    host = (parsed.hostname or '').strip().lower()
    if host.startswith('[') and host.endswith(']'):
        host = host[1:-1]
    if not host or not _hostname_allowed(host):
        return None, 'ollama_hote_non_autorise'
    return url.rstrip('/'), None


def build_ollama_chat_payload(
    model: str,
    messages: list[dict[str, str]],
    temperature: float = 0.3,
) -> dict[str, Any]:
    temp = max(0.0, min(2.0, float(temperature)))
    return {
        'model': model,
        'messages': messages,
        'stream': False,
        'options': {'temperature': temp},
    }


def call_ollama_chat(
    messages: list[dict[str, str]],
    temperature: float = 0.3,
) -> tuple[str | None, str | None]:
    """
    Appelle Ollama /api/chat. Retourne (réponse_texte, erreur_code).
    """
    if not getattr(settings, 'SOMA_AI_ENABLED', False):
        return None, 'soma_ai_desactive'

    base, err = validate_ollama_base_url(getattr(settings, 'OLLAMA_BASE_URL', 'http://127.0.0.1:11434'))
    if err:
        logger.warning('SOMA AI : URL Ollama refusée (%s)', err)
        return None, err

    model = getattr(settings, 'OLLAMA_MODEL', 'llama3.2')
    timeout = getattr(settings, 'OLLAMA_TIMEOUT_SEC', 120)
    url = base + '/api/chat'
    payload = build_ollama_chat_payload(model, messages, temperature=temperature)
    body = json.dumps(payload).encode('utf-8')
    req = urllib.request.Request(
        url,
        data=body,
        headers={'Content-Type': 'application/json'},
        method='POST',
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode('utf-8', errors='replace')
            data = json.loads(raw)
    except urllib.error.HTTPError as e:
        logger.warning('SOMA AI : HTTP Ollama %s', e.code)
        return None, 'ollama_http_error'
    except urllib.error.URLError as e:
        logger.warning('SOMA AI : connexion Ollama impossible (%s)', type(e).__name__)
        return None, 'ollama_connexion'
    except json.JSONDecodeError:
        return None, 'ollama_reponse_invalide'
    except Exception as e:
        logger.warning('SOMA AI : erreur appel Ollama (%s)', type(e).__name__)
        return None, 'ollama_erreur'

    msg = (data or {}).get('message') or {}
    content = msg.get('content')
    if not isinstance(content, str):
        return None, 'ollama_reponse_vide'
    content = _strip_control_chars(content.strip())
    max_out = getattr(settings, 'SOMA_AI_MAX_ASSISTANT_CHARS', 16000)
    if len(content) > max_out:
        content = content[:max_out].rstrip() + '\n\n[… réponse tronquée pour sécurité]'
    return content, None


def sanitize_history(session_msgs: list[dict[str, str]], max_msgs: int) -> list[dict[str, str]]:
    """Garde uniquement user/assistant alternés, taille bornée."""
    out: list[dict[str, str]] = []
    for m in session_msgs[-max_msgs:]:
        if not isinstance(m, dict):
            continue
        role = m.get('role')
        content = m.get('content')
        if role not in ('user', 'assistant') or not isinstance(content, str):
            continue
        content = _strip_control_chars(content.strip())
        if not content:
            continue
        max_each = getattr(settings, 'SOMA_AI_MAX_MESSAGE_CHARS', 4000)
        if len(content) > max_each:
            content = content[:max_each]
        out.append({'role': role, 'content': content})
    return out


SESSION_KEY = 'soma_ai_msgs'

# Dernière capture KPI « Bons de commande » (serveur) pour l’API chat — prioritaire sur le JSON client.
SOMA_SESSION_KPIS_ACHATS_BC = 'soma_session_kpis_achats_bons_commande'


def session_get_messages(request) -> list[dict[str, str]]:
    raw = request.session.get(SESSION_KEY)
    if not isinstance(raw, list):
        return []
    return sanitize_history(raw, getattr(settings, 'SOMA_AI_MAX_CONVERSATION_MESSAGES', 16))


def session_reset(request) -> None:
    request.session.pop(SESSION_KEY, None)
    request.session.modified = True


def sanitize_dashboard_context(raw: Any) -> str | None:
    """
    Valide un objet « contexte écran » envoyé par le client (structure simple, clés snake_case).
    Retourne une chaîne JSON compacte pour injection dans le prompt, ou None si invalide / vide.
    """
    if raw is None:
        return None
    if not isinstance(raw, dict):
        return None
    page = raw.get('page')
    if not isinstance(page, str) or not re.fullmatch(r'[a-z0-9_]{1,48}', page):
        page = 'unknown'
    kpis = raw.get('kpis')
    if not isinstance(kpis, dict):
        return None
    clean: dict[str, Any] = {}
    for k, v in list(kpis.items())[:24]:
        if not isinstance(k, str) or not re.fullmatch(r'[a-z0-9_]{1,32}', k):
            continue
        if isinstance(v, bool):
            clean[k] = v
        elif isinstance(v, int) and -10**12 < v < 10**12:
            clean[k] = v
        elif isinstance(v, float) and abs(v) < 10**15 and v == v:
            clean[k] = round(v, 4)
        elif isinstance(v, str) and len(v) <= 120:
            clean[k] = v[:120]
    if not clean:
        return None
    blob = {'page': page, 'kpis': clean}
    s = json.dumps(blob, ensure_ascii=False, separators=(',', ':'))
    if len(s) > 1200:
        return s[:1197] + '…'
    return s


def merge_session_achats_bc_kpis(request, dashboard_context: str | None) -> str | None:
    """
    Remplace les KPI « achats_bons_commande » par la dernière valeur en session (chargée avec la page),
    pour ignorer un contexte client altéré ou désynchronisé.
    """
    if not dashboard_context:
        return None
    try:
        d = json.loads(dashboard_context)
    except (json.JSONDecodeError, TypeError):
        return dashboard_context
    if not isinstance(d, dict) or d.get('page') != 'achats_bons_commande':
        return dashboard_context
    trusted = request.session.get(SOMA_SESSION_KPIS_ACHATS_BC)
    if not isinstance(trusted, dict) or trusted.get('page') != 'achats_bons_commande':
        return dashboard_context
    tk = trusted.get('kpis')
    if not isinstance(tk, dict):
        return dashboard_context
    merged = dict(d.get('kpis') or {})
    for key in (
        'nb_retard',
        'bons_en_retard',
        'total_bons',
        'montant_ttc_total',
        'nb_confirmes',
        'taux_confirmation_pct',
        'periode_debut',
        'periode_fin',
    ):
        if key in tk:
            merged[key] = tk[key]
    d['kpis'] = merged
    return sanitize_dashboard_context(d)


_RETARD_BONS_QUESTION_RE = re.compile(
    r'(?i)('
    r'bon.*retard|retard.*bon|bons?\s+de\s+commande.*retard|retard.*commande|'
    r'date\s*pr[ée]vue\s*d[ée]pass|en\s+retard'
    r')',
)


def try_direct_kpi_reply(user_message: str, dashboard_context: str | None) -> str | None:
    """
    Réponses factuelles sans LLM quand le contexte écran suffit (fiabilité des chiffres).
    """
    if not dashboard_context or not isinstance(user_message, str):
        return None
    t = user_message.strip().lower()
    if not t:
        return None
    if not _RETARD_BONS_QUESTION_RE.search(t):
        if 'retard' not in t or len(t) > 48:
            return None
    try:
        d = json.loads(dashboard_context)
    except (json.JSONDecodeError, TypeError):
        return None
    if not isinstance(d, dict) or d.get('page') != 'achats_bons_commande':
        return None
    k = d.get('kpis')
    if not isinstance(k, dict):
        return None
    nr = k.get('nb_retard')
    if nr is None:
        return None
    try:
        nr_i = int(nr)
    except (TypeError, ValueError):
        return None
    return (
        f'Sur l’écran « Bons de commande » (filtres actuels), le nombre de bons en retard '
        f'(date prévue dépassée, même logique que la carte KPI « Bons en retard ») est de {nr_i}.'
    )


def _dashboard_context_followup_fr(dashboard_context: str) -> str:
    """
    Complète le prompt avec des règles métier pour éviter les confusions fréquentes du modèle
    (ex. «bons en retard» ≠ total_bons - nb_retard).
    """
    try:
        d = json.loads(dashboard_context)
    except (json.JSONDecodeError, TypeError):
        return ''
    if not isinstance(d, dict) or d.get('page') != 'achats_bons_commande':
        return ''
    k = d.get('kpis')
    if not isinstance(k, dict):
        return ''
    tb = k.get('total_bons')
    nr = k.get('nb_retard')
    if tb is None or nr is None:
        return ''
    try:
        tb_i = int(tb)
        nr_i = int(nr)
    except (TypeError, ValueError):
        return ''
    faux = tb_i - nr_i
    return (
        '\n— Précision obligatoire pour cet écran «Bons de commande» : '
        '«bons en retard» / «date prévue dépassée» = uniquement nb_retard '
        f'(et bons_en_retard si présent), soit {nr_i}. '
        f'total_bons = {tb_i} (volume total avec les filtres actuels). '
        f'Le nombre {faux} (= {tb_i} - {nr_i}) n\'est pas le nombre de bons en retard ; '
        'ne le cite pas pour répondre à une question sur les retards.'
    )


def session_append(request, role: str, content: str) -> None:
    msgs = request.session.get(SESSION_KEY)
    if not isinstance(msgs, list):
        msgs = []
    msgs.append({'role': role, 'content': content})
    cap = getattr(settings, 'SOMA_AI_MAX_CONVERSATION_MESSAGES', 16)
    request.session[SESSION_KEY] = sanitize_history(msgs, cap)
    request.session.modified = True


def run_turn(
    request,
    user_message: str,
    dashboard_context: str | None = None,
) -> tuple[str | None, str | None]:
    """
    Enchaîne : rate limit → historique session → appel Ollama → mise à jour session.
    Retourne (réponse, erreur_code).

    ``dashboard_context`` : JSON court déjà validé (voir ``sanitize_dashboard_context``).
    """
    uid = request.user.pk
    if uid is None:
        return None, 'non_authentifie'

    if not _rate_limit_ok(int(uid)):
        return None, 'trop_de_requetes'

    max_chars = getattr(settings, 'SOMA_AI_MAX_MESSAGE_CHARS', 4000)
    clean, err = sanitize_user_message(user_message, max_chars)
    if err:
        return None, err

    logger.info(
        'SOMA AI : message utilisateur id=%s len=%s',
        uid,
        len(clean),
    )

    session_append(request, 'user', clean)

    direct = try_direct_kpi_reply(clean, dashboard_context)
    if direct:
        session_append(request, 'assistant', direct)
        return direct, None

    history = session_get_messages(request)

    messages_for_model: list[dict[str, str]] = [{'role': 'system', 'content': SYSTEM_PROMPT_FR}]
    if dashboard_context:
        followup = _dashboard_context_followup_fr(dashboard_context)
        messages_for_model.append({
            'role': 'system',
            'content': (
                'Indicateurs tableau de bord (valeurs calculées côté serveur, alignées sur l’écran affiché) : '
                + dashboard_context
                + followup
                + '\nPour toute question portant sur ces indicateurs, réutilise exactement ces chiffres. '
                  'N’extrapole pas au-delà. Hors périmètre de ce bloc, n’invente pas de données métier. '
                  'Réponse utilisateur : français uniquement.'
            ),
        })
    messages_for_model.extend(history)

    ollama_temp = 0.12 if dashboard_context else 0.3
    reply, err = call_ollama_chat(messages_for_model, temperature=ollama_temp)
    if err:
        session_pop_last_user_if_failed(request)
        return None, err

    session_append(request, 'assistant', reply)
    return reply, None


def session_pop_last_user_if_failed(request) -> None:
    """Si Ollama échoue, retire le dernier message user pour éviter un historique incohérent."""
    msgs = request.session.get(SESSION_KEY)
    if isinstance(msgs, list) and msgs and msgs[-1].get('role') == 'user':
        msgs.pop()
        request.session[SESSION_KEY] = msgs
        request.session.modified = True


def get_public_status(request) -> dict[str, Any]:
    enabled = getattr(settings, 'SOMA_AI_ENABLED', False)
    return {
        'status': 'ok',
        'soma_ai_enabled': enabled,
        'privacy': 'local_ollama' if enabled else 'disabled',
        'hint': (
            'Les échanges passent par un modèle hébergé localement (Ollama), sans envoi vers un cloud public.'
            if enabled else
            'SOMA AI est désactivé. Activez SOMA_AI_ENABLED et un serveur Ollama local pour l’utiliser.'
        ),
    }


ERROR_MESSAGES_FR = {
    'soma_ai_desactive': (
        'SOMA AI est désactivé sur ce serveur. Contactez l’administrateur pour activer Ollama en local '
        '(voir documentation SOMA_AI_* dans les paramètres).'
    ),
    'ollama_hote_non_autorise': (
        'Configuration refusée pour des raisons de sécurité : Ollama doit être joignable uniquement '
        'en local (127.0.0.1 / localhost), sauf exception LAN privée explicitement activée.'
    ),
    'ollama_connexion': 'Impossible de joindre le serveur Ollama local. Vérifiez qu’il est démarré (ollama serve).',
    'ollama_http_error': 'Le serveur Ollama a répondu avec une erreur. Vérifiez le nom du modèle (OLLAMA_MODEL).',
    'ollama_reponse_invalide': 'Réponse Ollama illisible.',
    'ollama_erreur': 'Erreur lors de l’appel au moteur IA local.',
    'ollama_reponse_vide': 'Le modèle n’a pas renvoyé de texte exploitable.',
    'message_vide': 'Veuillez saisir un message.',
    'message_trop_long': 'Message trop long. Réduisez le texte.',
    'message_invalide': 'Message invalide.',
    'trop_de_requetes': 'Trop de messages par minute. Patientez un instant.',
    'non_authentifie': 'Session expirée. Reconnectez-vous.',
}
