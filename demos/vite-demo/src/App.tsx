import { useState } from 'react'
import { build, showcase } from 'ts-pptx-demos-showcases/quarterly-review'
import { DECK_SOURCE } from './enums'
import logo from './assets/logo.png'
import './scss/styles.scss'

/**
 * The browser showcase.
 *
 * It builds the *same* deck as `pnpm demos:build quarterly-review` — same module, same
 * eleven slides — in the browser, with no server involved. That is the point of the demo:
 * the deck code is not written twice, and nothing in it knows which runtime it is on.
 *
 * The Field Notes showcase is deliberately not offered here. It loads photographs and a
 * video from disk by path, which a browser cannot do without those assets being served.
 */
function App() {
	const [state, setState] = useState<'idle' | 'building' | 'done' | 'error'>('idle')
	const [message, setMessage] = useState('')

	async function generate() {
		setState('building')
		setMessage('')
		try {
			// In the browser `writeFile` triggers a download rather than writing to disk,
			// so the file name is all the "path" this needs.
			await build(showcase.fileName)
			setState('done')
		} catch (err) {
			setState('error')
			setMessage(err instanceof Error ? err.message : String(err))
		}
	}

	const htmlNav = () => (
		<nav className="navbar navbar-expand-lg bg-primary" data-bs-theme="dark">
			<div className="container-fluid">
				<a className="navbar-brand" href="https://github.com/shbernal/ts-pptx">
					<img src={logo} alt="logo" width="30" height="30" className="d-inline-block align-text-center me-2" />
					ts-pptx
				</a>
				<div className="hstack gap-2 ms-auto">
					<a className="btn btn-outline-light btn-sm" href="https://github.com/shbernal/ts-pptx#readme">
						Docs
					</a>
					<a className="btn btn-outline-light btn-sm" href="https://github.com/shbernal/ts-pptx/releases">
						Releases
					</a>
					<a className="btn btn-outline-light btn-sm" href="https://github.com/shbernal/ts-pptx">
						GitHub
					</a>
				</div>
			</div>
		</nav>
	)

	const htmlStatus = () => {
		if (state === 'done') {
			return (
				<div className="alert alert-success mb-0" role="status">
					Built <code>{showcase.fileName}</code> — check your downloads.
				</div>
			)
		}
		if (state === 'error') {
			return (
				<div className="alert alert-danger mb-0" role="alert">
					{message}
				</div>
			)
		}
		return null
	}

	const htmlMain = () => (
		<main className="container my-5">
			<div className="card">
				<div className="card-header">
					<h1 className="display-5">{showcase.title}</h1>
					<div className="lead text-primary-emphasis">{showcase.description}</div>
				</div>
				<div className="card-body">
					<p>
						This React + TypeScript + Vite app imports the showcase deck module directly and builds it in the
						browser. No server, no round trip — the <code>.pptx</code> is assembled in the tab and handed to
						the download manager.
					</p>
					<h5 className="text-info mt-4">How the page calls it</h5>
					<pre className="bg-black mt-3 p-3 rounded">
						<code className="language-typescript" style={{ fontSize: '0.8rem' }}>
							{DECK_SOURCE}
						</code>
					</pre>
				</div>
				<div className="card-footer p-3 d-grid gap-3">
					<button type="button" className="btn btn-success btn-lg" onClick={generate} disabled={state === 'building'}>
						{state === 'building' ? 'Building…' : `Build ${showcase.fileName}`}
					</button>
					{htmlStatus()}
				</div>
			</div>
		</main>
	)

	return (
		<section>
			{htmlNav()}
			{htmlMain()}
		</section>
	)
}

export default App
