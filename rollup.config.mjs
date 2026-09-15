import typescript from '@rollup/plugin-typescript';

export default {
  input: 'src/index.ts',
  output: {
    file: 'dist/Code.js',
    format: 'es',
    banner: '// Generated from TypeScript - do not edit directly\n',
  },
  // Disable tree-shaking so Apps Script entry points are preserved
  treeshake: false,
  plugins: [
    typescript({
      tsconfig: './tsconfig.json',
      compilerOptions: {
        noEmit: false,
        declaration: false,
      },
    }),
  ],
  external: [],
};
