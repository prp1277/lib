import React from "react";

const Card = () => {
  return (
    <div className="row">
      <div class="col s12 m4">
        <div className="card blue-grey lighten-1">
          <div className="card-image">
            <img src="https://lorempixel.com/100/190/nature/6" alt="img" />
          </div>
          <div className="card-content white-text">
            <span className="card-title">Card Title</span>
            <p>Card Content</p>
          </div>
          <div class="card-action">
            <a href="/">Link To Article</a>
          </div>
        </div>
      </div>
    </div>
  );
};

export default Card;
