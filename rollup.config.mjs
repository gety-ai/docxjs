import typescript from '@rollup/plugin-typescript';
import terser from '@rollup/plugin-terser';
import commonjs from '@rollup/plugin-commonjs';
import nodeResolve from '@rollup/plugin-node-resolve';
import replace from '@rollup/plugin-replace';

const output = {
	banner: `/*
 * @license
 * docx-preview <https://github.com/VolodymyrBaydalka/docxjs>
 * Released under Apache License 2.0  <https://github.com/VolodymyrBaydalka/docxjs/blob/master/LICENSE>
 * Copyright Volodymyr Baydalka
 */`,
	sourcemap: true,
}

const umdOutput = {
	...output,
	name: "docx",
	file: 'dist/docx-preview.js',
	format: 'umd',
};

const workerOutput = {
	...output,
	name: "docxWorker",
	file: 'dist/docx-preview-worker.js',
	format: 'iife',
};

function createPlugins() {
	return [
		replace({
			preventAssignment: true,
			'process.env.NODE_ENV': JSON.stringify('production')
		}),
		nodeResolve(),
		commonjs(),
		typescript()
	];
}

export default args => {
	const mainConfig = {
		input: 'src/docx-preview.ts',
		output: [umdOutput],
		plugins: createPlugins()
	}

	const workerConfig = {
		input: 'src/parser-worker.ts',
		output: [workerOutput],
		plugins: createPlugins()
	}

	if (args.environment == 'BUILD:production')
		mainConfig.output = [umdOutput,
			{
				...umdOutput,
				file: 'dist/docx-preview.min.js',
				plugins: [terser()]
			},
			{
				...output,
				file: 'dist/docx-preview.mjs',
				format: 'es',
			},
			{
				...output,
				file: 'dist/docx-preview.min.mjs',
				format: 'es',
				plugins: [terser()]
			}];

	return [mainConfig, workerConfig]
};
