import { aiFetch } from './fetch'
import { getProviderAdapter } from './registry'
import type { AiProviderConfig, AiProviderId, FetchModelsResult, RemoteModel } from './types'

/** OpenRouter's own richer /models schema (distinct from the generic OpenAI-compatible list) */
interface OpenRouterModelsResponse {
  data?: Array<{ id?: string; name?: string; context_length?: number }>
}

/** the generic OpenAI-compatible GET /models shape Ollama's compat layer returns */
interface OpenAiModelsResponse {
  data?: Array<{ id?: string }>
}

function authHeaders(apiKey: string): Record<string, string> {
  return apiKey ? { Authorization: `Bearer ${apiKey}` } : {}
}

async function fetchOpenRouterModels(config: AiProviderConfig): Promise<FetchModelsResult> {
  const response = await aiFetch('https://openrouter.ai/api/v1/models', {
    headers: authHeaders(config.apiKey),
  })
  if (!response.ok) return { ok: false, models: [], error: `HTTP ${response.status}` }
  const json = (await response.json()) as OpenRouterModelsResponse
  const models: RemoteModel[] = (json.data ?? [])
    .filter((m): m is { id: string; name?: string; context_length?: number } => !!m.id)
    .map((m) => ({
      id: m.id,
      ...(m.name !== undefined ? { label: m.name } : {}),
      ...(m.context_length !== undefined ? { contextLength: m.context_length } : {}),
    }))
  return { ok: true, models }
}

/** local Ollama and Ollama Cloud both expose the OpenAI-compatible GET /models route */
async function fetchOllamaModels(
  provider: 'ollama' | 'ollamaCloud',
  config: AiProviderConfig,
): Promise<FetchModelsResult> {
  const baseUrl = config.baseUrl || getProviderAdapter(provider).resolveEndpoint(config).baseUrl
  const response = await aiFetch(`${baseUrl.replace(/\/$/, '')}/models`, {
    headers: authHeaders(config.apiKey),
  })
  if (!response.ok) return { ok: false, models: [], error: `HTTP ${response.status}` }
  const json = (await response.json()) as OpenAiModelsResponse
  const models: RemoteModel[] = (json.data ?? [])
    .filter((m): m is { id: string } => !!m.id)
    .map((m) => ({ id: m.id }))
  return { ok: true, models }
}

/**
 * Live model catalog for a provider whose meta declares dynamicModels — OpenRouter's
 * full model list, or whatever's installed/available on a reachable Ollama server.
 * Network/parse failures resolve to ok:false rather than throwing, so a settings-pane
 * fetch degrades to the existing free-text model field instead of crashing the pane.
 */
export async function fetchProviderModels(
  provider: AiProviderId,
  config: AiProviderConfig,
): Promise<FetchModelsResult> {
  try {
    if (provider === 'openrouter') return await fetchOpenRouterModels(config)
    if (provider === 'ollama' || provider === 'ollamaCloud') {
      return await fetchOllamaModels(provider, config)
    }
    return { ok: false, models: [], error: `${provider} has no live model catalog` }
  } catch (error) {
    return { ok: false, models: [], error: error instanceof Error ? error.message : String(error) }
  }
}
