import { describe, expect, it } from 'vitest'
import { createElectronBuilderConfig } from '../scripts/electron-builder-config.mjs'

const unsignedWindows = {
  DSH_DESKTOP_APP_ID: 'com.example.installer',
  DSH_DESKTOP_MANDATORY_UPDATE_TEST_ORIGIN: 'https://policy.example.com',
  DSH_DESKTOP_MANDATORY_UPDATE_CONFIG: JSON.stringify({ allowedAuthOrigins: ['https://login.example.com'] }),
  DSH_DESKTOP_TARGET_PLATFORM: 'win32',
  DSH_DESKTOP_TARGET_ARCH: 'x64',
  DSH_DESKTOP_UNSIGNED: '1',
}
const feed = 'https://github.com/example/dsh-win7/releases/latest/download/'

describe('Windows 7 update feed', () => {
  it('leaves an unsigned build without a feed by default', () => {
    expect(createElectronBuilderConfig(unsignedWindows, 'win32', 'x64').publish).toBeNull()
  })

  it('publishes an unsigned Windows build to DSH_WIN7_UPDATE_URL one byte range at a time', () => {
    const config = createElectronBuilderConfig({ ...unsignedWindows, DSH_WIN7_UPDATE_URL: feed }, 'win32', 'x64')
    expect(config.publish).toEqual([{ provider: 'generic', url: feed, channel: 'nightly', useMultipleRangeRequest: false }])
    // The declared config type lists only the app ID among the extra metadata fields.
    expect((config.extraMetadata as Record<string, unknown>).dshDesktopUpdateChannel).toBe('nightly')
  })

  it('rejects a feed that is not an http(s) directory', () => {
    expect(() => createElectronBuilderConfig({ ...unsignedWindows, DSH_WIN7_UPDATE_URL: feed.slice(0, -1) }, 'win32', 'x64'))
      .toThrow('DSH_WIN7_UPDATE_URL')
    expect(() => createElectronBuilderConfig({ ...unsignedWindows, DSH_WIN7_UPDATE_URL: 'file:///C:/feed/' }, 'win32', 'x64'))
      .toThrow('DSH_WIN7_UPDATE_URL')
  })
})
