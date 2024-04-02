const typescript = require('@rollup/plugin-typescript');
const peerDepsExternal = require('rollup-plugin-peer-deps-external');
const resolve = require('@rollup/plugin-node-resolve').default;
const commonjs = require('@rollup/plugin-commonjs');
const nodePolyfills = require('rollup-plugin-node-polyfills');

module.exports = {
  input: 'src/index.ts',
  output: {
    dir: 'dist',
    format: 'cjs',
    sourcemap: true,
  },
  plugins: [
    peerDepsExternal(),
    resolve(),
    typescript(),
    commonjs(),
    nodePolyfills()
  ],
};
