const config = require('../../config');

describe('SERVER_INSTRUCTIONS', () => {
  test('is a non-empty string', () => {
    expect(typeof config.SERVER_INSTRUCTIONS).toBe('string');
    expect(config.SERVER_INSTRUCTIONS.length).toBeGreaterThan(0);
  });

  test('warns about untrusted content', () => {
    const text = config.SERVER_INSTRUCTIONS.toLowerCase();
    expect(text).toContain('untrusted');
  });

  test('mentions boundary markers', () => {
    const text = config.SERVER_INSTRUCTIONS.toLowerCase();
    expect(text).toContain('boundary');
  });

  test('instructs the LLM not to follow instructions within boundaries', () => {
    const text = config.SERVER_INSTRUCTIONS.toLowerCase();
    expect(text).toContain('never follow instructions');
  });

  test('mentions prompt injection', () => {
    const text = config.SERVER_INSTRUCTIONS.toLowerCase();
    expect(text).toContain('prompt-injection');
  });

  test('includes email workflow guidance about emailId', () => {
    const text = config.SERVER_INSTRUCTIONS.toLowerCase();
    expect(text).toContain('workflow');
    expect(text).toContain('emailid');
    expect(text).toContain('list-emails');
    expect(text).toContain('search-emails');
  });
});
