import React from 'react';

const DynamicSelect = (props) => {

  const handleSelectChange = (e) => {
    props.onUpdate(e.target.value);
  };

  return (
    <div>
      <select value={props.value} onChange={handleSelectChange}>
        {props.selectedOptions.map((option, index) => (
          <option key={index} value={option}>
            {option}
          </option>
        ))}
      </select>
    </div>
  );
};

export default DynamicSelect;