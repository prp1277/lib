// credit to developer-cheatsheets.com: http://www.developer-cheatsheets.com/react

/**----------------------------------------------------------------*
* Create A Stateless Component
* -----------------------------------------------------------------*
*/
import React from 'react'

const YourComponent = () => <div>aaa</div>

export default YourComponent

/**----------------------------------------------------------------*
* Create a Class Component 
* -----------------------------------------------------------------*
*/
import React from 'react'

class YourComponent extends React.Component {
  render() {
    return <div>aaa</div>
  }
}

export default YourComponent

/**----------------------------------------------------------------*
* Properties in a Stateless Component
* -----------------------------------------------------------------*
*/
const YourComponent = ({ propExample1, example2 }) => (
  <div>
    <h1>properties from parent component:</h1>
    <ul>
      <li>{propExample1}</li>
      <li>{example2}</li>
    </ul>
  </div>
)

// <YourComponent propExample1="aaa" example2="bbb" />

/**----------------------------------------------------------------*
* Properties in a Class Component
* -----------------------------------------------------------------*
*/
class YourComponent extends React.Component {
  render() {
    return (
      <div>
        <h1>
          properties from parent component:
        </h1>
        <ul>
          <li>{this.props.propExample1}</li>
          <li>{this.props.example2}</li>
        </ul>
      </div>
    )
  }
}

/**----------------------------------------------------------------*
* Handling Children
* -----------------------------------------------------------------*
*/
const Component1 = (props) => (
  <div>{props.children}</div>
)

const Component2 = () => (
  <Component1>
    <h1>Component 1</h1>
  </Component1>
)

/**----------------------------------------------------------------*
* State
* -----------------------------------------------------------------*
*/
class CountClicks extends React.Component {
  state = {
    clicks: 0
  }

  onButtonClick = () => {
    this.setState(prevState => ({
      clicks: prevState.clicks + 1
    }))
  }

  render() {
    return (
      <div>
        <button onClick={this.onButtonClick}>
          Click me
        </button>
        <span>
          The button clicked 
          {this.state.clicks} times.
        </span>
      </div>
    )
  }
}

/**----------------------------------------------------------------*
* React Router 
* Cheatsheet: http://www.developer-cheatsheets.com/react-router
* -----------------------------------------------------------------*
*/
import { 
  BrowserRouter, 
  Route 
} from 'react-router-dom'

const Hello = () => <h1>Hello world!</h1>

const App = () => (
  <BrowserRouter>
    <div>
      <Route path="/hello" component={Hello}/>
    </div>
  </BrowserRouter>
)

// open: http://localhost:3000/hello

/**----------------------------------------------------------------*
* React Redux Provider
* -----------------------------------------------------------------*
*/
import React from 'react'
import { render } from 'react-dom'
import { Provider } from 'react-redux'
import { createStore } from 'redux'
import todoApp from './reducers'
import App from './components/App'
​
const store = createStore(todoApp)
​
render(
  <Provider store={store}>
    <App />
  </Provider>,
  document.getElementById('root')
)

/**----------------------------------------------------------------*
* React Redux connect
* -----------------------------------------------------------------*
*/
import { connect } from 'react-redux'
​
YourComponent = connect(
  mapStateToProps,
  mapDispatchToProps
)(YourComponent)
​
export default YourComponent

/**----------------------------------------------------------------*
* Ref
* -----------------------------------------------------------------*
*/
class AutoFocusTextInput extends React.Component {
  constructor(props) {
    super(props);
    this.textInput = React.createRef();
  }

  componentDidMount() {
    this.textInput.current.focus();
  }

  render() {
    return (
      <input ref={this.textInput} />
    );
  }
}

/**----------------------------------------------------------------*
* Higher Order Components
* -----------------------------------------------------------------*
*/
import React from 'react'
import Loading from '../components/Loading'

const withLoading = WrappedComponent => {
  return (props = {}) => {
    if (props.isLoading) {
      return <Loading />
    }
    return <WrappedComponent {...this.props} />
  }
}

export default withLoading

// --------

const MyComponent = ({}) => <div /> // ...
const WithLoadingComponent = withLoading(MyComponent)