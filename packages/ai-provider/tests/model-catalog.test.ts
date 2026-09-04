import { afterEach, describe, expect, it, vi } from 'vitest'
import { fetchProviderModels } from '../src/model-catalog'
import { errorResponse, jsonResponse } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('fetchProviderModels', () => {
  it('openrouter: maps the catalog response to RemoteModel, auth header optional', async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      jsonResponse({
        data: [
          { id: 'anthropic/claude-sonnet-5', name: 'Claude Sonnet 5', context_length: 200000 },
          { id: 'openrouter/auto' },
        ],
      }),
    )
    vi.stubGlobal('fetch', fetchMock)
    const result = await fetchProviderModels('openrouter', { apiKey: '', model: '' })
    expect(result).toEqual({
      ok: true,
      models: [
        { id: 'anthropic/claude-sonnet-5', label: 'Claude Sonnet 5', contextLength: 200000 },
        { id: 'openrouter/auto', label: undefined, contextLength: undefined },
      ],
    })
    const [url, init] = fetchMock.mock.calls[0]
    expect(url).toBe('https://openrouter.ai/api/v1/models')
    expect(init.headers).toEqual({})
  })

  it('openrouter: sends a bearer header when a key is configured', async () => {
    const fetchMock = vi.fn().mockResolvedValue(jsonResponse({ data: [] }))
    vi.stubGlobal('fetch', fetchMock)
    await fetchProviderModels('openrouter', { apiKey: 'sk-or-x', model: '' })
    expect(fetchMock.mock.calls[0][1].headers).toEqual({ Authorization: 'Bearer sk-or-x' })
  })

  it('ollama: hits the configured base URL, defaulting to the local server', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ data: [{ id: 'llama3.1' }, { id: 'qwen2.5-coder' }] }))
    vi.stubGlobal('fetch', fetchMock)
    const result = await fetchProviderModels('ollama', { apiKey: '', model: '' })
    expect(result).toEqual({ ok: true, models: [{ id: 'llama3.1' }, { id: 'qwen2.5-coder' }] })
    expect(fetchMock.mock.calls[0][0]).toBe('http://localhost:11434/v1/models')
  })

  it('ollama: a stored base URL overrides the local default', async () => {
    const fetchMock = vi.fn().mockResolvedValue(jsonResponse({ data: [] }))
    vi.stubGlobal('fetch', fetchMock)
    await fetchProviderModels('ollama', { apiKey: '', model: '', baseUrl: 'http://192.168.1.5:11434/v1' })
    expect(fetchMock.mock.calls[0][0]).toBe('http://192.168.1.5:11434/v1/models')
  })

  it('ollamaCloud: hits ollama.com with the bearer key', async () => {
    const fetchMock = vi.fn().mockResolvedValue(jsonResponse({ data: [{ id: 'gpt-oss:120b' }] }))
    vi.stubGlobal('fetch', fetchMock)
    const result = await fetchProviderModels('ollamaCloud', { apiKey: 'sk-cloud', model: '' })
    expect(result).toEqual({ ok: true, models: [{ id: 'gpt-oss:120b' }] })
    const [url, init] = fetchMock.mock.calls[0]
    expect(url).toBe('https://ollama.com/v1/models')
    expect(init.headers).toEqual({ Authorization: 'Bearer sk-cloud' })
  })

  it('resolves ok:false (not a throw) on an HTTP error', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(errorResponse(401, 'unauthorized')))
    const result = await fetchProviderModels('ollamaCloud', { apiKey: 'bad', model: '' })
    expect(result).toEqual({ ok: false, models: [], error: 'HTTP 401' })
  })

  it('resolves ok:false on a network failure', async () => {
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('ECONNREFUSED')))
    const result = await fetchProviderModels('ollama', { apiKey: '', model: '' })
    expect(result).toEqual({ ok: false, models: [], error: 'ECONNREFUSED' })
  })

  it('rejects a provider with no live catalog instead of guessing an endpoint', async () => {
    const result = await fetchProviderModels('anthropic', { apiKey: 'k', model: 'claude-sonnet-5' })
    expect(result).toEqual({ ok: false, models: [], error: 'anthropic has no live model catalog' })
  })
})
