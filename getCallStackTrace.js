function getCallStackTrace(ignoreAnonymous = true) {
  const error = new Error();
  const stack = error.stack.split('\n').slice(2); // remove "Error" and self

  const functionNames = stack
    .map(line => {
      const match = line.match(/at ([\w$]+) /);
      return match ? match[1] : 'anonymous';
    })
    .filter(name => !(ignoreAnonymous && name === 'anonymous'));

  return functionNames.reverse().join(' → ');
}